#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# barcodeSuite.py - BarcodeSuite: FASTQ Inspector, FASTA Tools, FASTA Compare, BLAST,
# GenBank Batch, BOLD Formatter

from __future__ import annotations

### Verificación e instalación automática de dependencias ###################
# Solo aplica cuando se corre desde código fuente (`python barcodeSuite.py`): el
# .exe empaquetado con PyInstaller ya trae todo embebido y no tiene pip ni
# un intérprete de Python independiente al que instalarle nada.
import sys
import subprocess
import importlib.util


def _install_package(package):
    """Verifica si un paquete está instalado. Si no, intenta instalarlo con pip."""
    if importlib.util.find_spec(package) is None:
        print(f"⚠️ {package} no encontrado. Intentando instalarlo...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", package], check=True)
            print(f"✅ {package} instalado correctamente.")
        except subprocess.CalledProcessError as e:
            print(f"❌ No se pudo instalar {package}: {e}")
            sys.exit(1)


def _check_and_install_pip():
    """Verifica si pip está instalado y lo instala si es necesario."""
    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "--version"],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except subprocess.CalledProcessError:
        print("⚠️ pip no encontrado. Intentando instalar...")
        try:
            if sys.platform == "win32":
                subprocess.run([sys.executable, "-m", "ensurepip"], check=True)
                subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], check=True)
            elif sys.platform.startswith("linux"):
                import shutil as _shutil
                package_manager = _shutil.which("apt") or _shutil.which("dnf") or _shutil.which("yum")
                if package_manager:
                    subprocess.run(["sudo", package_manager, "install", "-y", "python3-pip"], check=True)
                else:
                    print("❌ No se pudo determinar el gestor de paquetes en Linux.")
                    sys.exit(1)
            else:
                print(f"❌ Sistema operativo no soportado: {sys.platform}")
                sys.exit(1)
            print("✅ pip instalado correctamente.")
        except subprocess.CalledProcessError as e:
            print(f"❌ No se pudo instalar pip: {e}")
            sys.exit(1)


if not getattr(sys, "frozen", False):
    _REQUIRED_PACKAGES = ["PyQt5", "edlib", "xlsxwriter", "openpyxl"]
    _check_and_install_pip()
    for _pkg in _REQUIRED_PACKAGES:
        _install_package(_pkg)

import os
import datetime
import time
import shutil
import fnmatch
import csv
import warnings
import multiprocessing
import threading
import subprocess
import xlsxwriter
import openpyxl
import itertools
import re
import edlib
from collections import Counter, OrderedDict
from typing import Dict, List, Optional, Tuple
from PyQt5 import QtCore

# Suppress non-relevant warnings from Biopython (partial codons, etc.)
warnings.filterwarnings("ignore", message="Partial codon")
warnings.filterwarnings("ignore", category=UserWarning, module="Bio")

from PyQt5 import QtCore, QtGui, QtWidgets

def _get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _profiles_dir() -> str:
    path = os.path.join(_get_base_dir(), "_profiles")
    os.makedirs(path, exist_ok=True)
    return path



# ── i18n ──────────────────────────────────────────────────────────────────────
import json as _json_mod
import xml.etree.ElementTree as _ET

_LANG_CONFIG = os.path.join(os.path.expanduser("~"), ".ontbarcoder_lang.json")

def _load_lang():
    try:
        with open(_LANG_CONFIG) as f:
            return _json_mod.load(f).get("lang", "en")
    except Exception:
        return "en"

def _save_lang(lang):
    try:
        with open(_LANG_CONFIG, "w") as f:
            _json_mod.dump({"lang": lang}, f)
    except Exception:
        pass

_CURRENT_LANG = [_load_lang()]
_TRANSLATIONS = {}

def _parse_ts(path):
    result = {}
    try:
        root = _ET.parse(path).getroot()
        for ctx in root.findall("context"):
            name = ctx.findtext("name", "")
            d = {}
            for msg in ctx.findall("message"):
                src = msg.findtext("source", "")
                tr_el = msg.find("translation")
                if tr_el is not None and tr_el.text and tr_el.get("type") != "unfinished":
                    d[src] = tr_el.text
            result[name] = d
    except Exception:
        pass
    return result

def _load_translations():
    ts_dir = os.path.join(_get_base_dir(), "translations")
    for lang in ["en"]:
        ts_path = os.path.join(ts_dir, f"{lang}.ts")
        if os.path.exists(ts_path):
            _TRANSLATIONS[lang] = _parse_ts(ts_path)

_load_translations()

def _tr(context, source):
    lang = _CURRENT_LANG[0]
    if lang == "es":
        return source
    return _TRANSLATIONS.get(lang, {}).get(context, {}).get(source, source)

def set_language(lang, panels=None):
    _CURRENT_LANG[0] = lang
    _save_lang(lang)
    if panels:
        for panel in panels:
            QtWidgets.QApplication.postEvent(
                panel, QtCore.QEvent(QtCore.QEvent.LanguageChange)
            )


# ── Color palette ──────────────────────────────────────────────────────
BLUE = "#185FA5"
BLUE_LIGHT = "#E6F1FB"
BLUE_MID = "#378ADD"
GREEN = "#3B6D11"
GREEN_LT = "#EAF3DE"
GREEN_MID = "#639922"
AMBER = "#854F0B"
AMBER_LT = "#FAEEDA"
RED = "#A32D2D"
WHITE = "#FFFFFF"
RED_LT = "#FCEBEB"
GRAY_BG = "#F5F5F3"
GRAY_CARD = "#F5F5F3"
GRAY_LINE = "#E0DED8"
TEXT_PRI = "#1A1A18"
TEXT_SEC = "#6B6960"
TEXT_HINT = "#A09D96"
SIDEBAR_BG = "#EDEDED"
TOPBAR_BG = "#2A2D3A"
GRAY_DARK = "#203864"

STYLESHEET = f"""
QWidget {{
    font-family: -apple-system, 'Segoe UI', Arial, sans-serif;
    font-size: 18px;
    color: {TEXT_PRI};
    background-color: transparent;
}}
QMainWindow, #root_bg {{
    background-color: {GRAY_BG};
}}

/* ── Dialogs — explicit white background to avoid inheritance of OS dark theme ── */
QDialog {{
    background-color: {GRAY_CARD};
    color: {TEXT_PRI};
}}
QDialog QLabel {{
    color: {TEXT_PRI};
    background-color: transparent;
}}
QDialog QRadioButton {{
    color: {TEXT_PRI};
    background-color: transparent;
}}
QDialog QPushButton {{
    color: {TEXT_PRI};
}}
QMessageBox {{
    background-color: {GRAY_CARD};
    color: {TEXT_PRI};
}}
QMessageBox QLabel {{
    color: {TEXT_PRI};
    background-color: transparent;
}}

/* ── Sidebar ── */
#sidebar {{
    background-color: {SIDEBAR_BG};
    border-right: 1px solid {GRAY_LINE};
}}
#sidebar_item {{
    padding: 12px 16px;
    border-left: 8px solid transparent;
    color: {TEXT_SEC};
    background: transparent;
    text-align: left;
    border-radius: 0;
    font-size: 20px;
}}
#sidebar_item:hover {{
    background-color: {GRAY_BG};
}}
#sidebar_item[state="active"] {{
    color: {GRAY_DARK};
    background-color: {GRAY_BG};
    border-left-color: {BLUE_MID};
    font-weight: 500;
}}
#sidebar_item[state="done"] {{
    color: {GREEN};
    background: transparent;
    border-left: 2px solid transparent;
}}
#sidebar_item[state="locked"] {{
    color: {TEXT_HINT};
    background: transparent;
    border-left: 2px solid transparent;
}}
#sidebar_section {{
    font-size: 10px;
    font-weight: 600;
    color: {TEXT_HINT};
    padding: 8px 16px 2px;
    letter-spacing: 0.5px;
    background: transparent;
}}

/* ── Topbar ── */
#topbar {{
    background-color: {TOPBAR_BG};
    border-bottom: 1px solid {GRAY_LINE};
}}
#topbar_logo {{
    font-size: 28px;
    font-weight: 600;
    letter-spacing: -0.3px;
    color: {WHITE};
}}
#topbar_badge {{
    font-size: 15px;
    padding: 2px 8px;
    border-radius: 10px;
    background-color: {GRAY_BG};
    color: {BLUE};
    margin-top: 8px;      /* ← Space above */
    margin-bottom: 8px;   /* ← Space below */
}}

/* ── Cards ── */
#card {{
    background-color: {GRAY_CARD};
    border: 1px solid {GRAY_LINE};
    border-radius: 10px;
    padding: 16px;
}}
#stat_card {{
    background-color: {GRAY_BG};
    border-radius: 8px;
    padding: 12px;
}}

/* ── Mode cards ── */
#mode_card {{
    background-color: {GRAY_CARD};
    border: 1px solid {GRAY_LINE};
    border-radius: 10px;
    padding: 16px;
}}
#mode_card:hover {{
    border-color: {BLUE_MID};
}}
#mode_card[selected="true"] {{
    border: 2px solid {BLUE_MID};
    background-color: {BLUE_LIGHT};
}}

/* ── Drop zones ── */
#drop_zone {{
    background-color: {GRAY_BG};
    border: 0.5px solid #E6E6E3;
    border-radius: 10px;
    padding: 24px;
}}
#drop_zone[dragging="true"] {{
    background-color: #EBEBEA;
    border: 2px dashed {BLUE_MID};
    border-radius: 10px;
    padding: 24px;
}}
#drop_zone[filled="true"] {{
    background-color: {GREEN_LT};
    border: 1px solid {GREEN_MID};
    border-style: solid;
}}

/* ── Buttons ── */
#primary_btn {{
    background-color: {BLUE};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 9px 20px;
    font-size: 18px;
    font-weight: 500;
}}
#primary_btn:hover {{ background-color: #0C4A82; }}
#primary_btn:pressed {{ background-color: #083460; }}
#primary_btn:disabled {{ background-color: {GRAY_LINE}; color: {TEXT_HINT}; }}

/* Unified style for secondary buttons (Search, Gold, MinKNOW, etc.) */
.secondary-btn, #secondary_btn, .basecalling-btn {{
    background-color: transparent;
    color: {BLUE};
    border: 1px solid {BLUE_MID};
    border-radius: 8px;
    padding: 7px 16px;
    font-size: 18px;
}}
.secondary-btn:hover, #secondary_btn:hover, .basecalling-btn:hover {{
    background-color: {BLUE_LIGHT};
    color: {BLUE};
    border-color: {BLUE};
}}
.secondary-btn:pressed, #secondary_btn:pressed, .basecalling-btn:pressed {{
    background-color: {BLUE_MID};
    color: white;
}}

#danger_btn {{
    background-color: transparent;
    color: {RED};
    border: 1px solid #F09595;
    border-radius: 8px;
    padding: 7px 14px;
    font-size: 18px;
}}
#danger_btn:hover {{
    background-color: {RED_LT};
    color: {RED};
    border-color: {RED};
}}
#danger_btn:pressed {{
    background-color: {RED};
    color: white;
}}

/* ── Inputs ── */
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {{
    background-color: {GRAY_CARD};
    border: 1px solid {GRAY_LINE};
    border-radius: 6px;
    padding: 5px 9px;
    font-size: 18px;
    selection-background-color: {BLUE_LIGHT};
    selection-color: {TEXT_PRI};
}}
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {{
    border-color: {BLUE_MID};
    background-color: {GRAY_CARD};
    color: {TEXT_PRI};
}}

QSpinBox QLineEdit, QDoubleSpinBox QLineEdit {{
    background-color: {GRAY_CARD};
    selection-background-color: {BLUE_LIGHT};
    selection-color: {TEXT_PRI};
    color: {TEXT_PRI};
}}
QComboBox {{ padding-right: 20px; }}
QComboBox::drop-down {{ width: 20px; subcontrol-origin: padding; subcontrol-position: top right; }}

/* ── Tabs ── */
QTabWidget {{
    background-color: {TOPBAR_BG};
}}
QTabBar {{
    background-color: {TOPBAR_BG};
}}
QTabWidget::pane {{
    border: 1px solid {GRAY_LINE};
    border-top: 3px solid {TOPBAR_BG};
    border-radius: 0 0 8px 8px;
    background: {GRAY_CARD};
}}
QTabBar::tab {{
    background: transparent;
    border-bottom: 5px solid transparent;
    padding: 7px 14px;
    font-size: 20px;
    color: #A8C4E8;
    margin-right: 2px;
    min-width: 200px;
}}
QTabBar::tab:selected {{
    color: {WHITE};
    border-bottom-color: {BLUE_MID};
    font-weight: 600;
}}
QTabBar::tab:hover:!selected {{
    color: {WHITE};
}}

QTabBar::tab-bar {{
    alignment: center;
}}

/* ── Progress bar ── */
QProgressBar {{
    background-color: {GRAY_LINE};
    border: none;
    border-radius: 3px;
    height: 6px;
    text-align: center;
}}
QProgressBar::chunk {{
    background-color: {BLUE_MID};
    border-radius: 3px;
}}

/* ── Log / text areas ── */
#log_area {{
    background-color: {GRAY_CARD};
    border: 1px solid {GRAY_LINE};
    border-radius: 8px;
    font-family: 'Courier New', monospace;
    font-size: 14px;
    color: {TEXT_SEC};
    padding: 10px;
}}

/* ── Scrollbars ── */
QScrollBar:vertical {{
    width: 6px;
    background: transparent;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {GRAY_LINE};
    border-radius: 3px;
    min-height: 24px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}

/* ── Phase rows ── */
#phase_row {{
    background-color: {GRAY_CARD};
    border: 1px solid {GRAY_LINE};
    border-radius: 8px;
    padding: 18px 14px;
}}
#phase_row[state="current"] {{
    border-color: {BLUE_MID};
    background-color: {BLUE_LIGHT};
}}
#phase_row[state="done"] {{
    border-color: {GREEN_MID};
    background-color: {GREEN_LT};
}}

/* ── Dialogs and buttons ── */
QDialog QDialogButtonBox QPushButton,
QMessageBox QPushButton {{
    background-color: transparent;
    color: #185FA5;
    border: 1px solid #378ADD;
    border-radius: 8px;
    padding: 7px 16px;
    font-size: 18px;
    min-width: 80px;
}}

QDialog QDialogButtonBox QPushButton:hover,
QMessageBox QPushButton:hover {{
    background-color: #E6F1FB;
    color: #185FA5;
    border-color: #185FA5;
}}

QDialog QDialogButtonBox QPushButton:pressed,
QMessageBox QPushButton:pressed {{
    background-color: #378ADD;
    color: white;
}}

/* Default button (the one with focus) */
QDialog QDialogButtonBox QPushButton[default="true"],
QMessageBox QPushButton[default="true"] {{
    background-color: #185FA5;
    color: white;
    border: none;
}}

QDialog QDialogButtonBox QPushButton[default="true"]:hover {{
    background-color: #0C4A82;
}}

"""


# ═══════════════════════════════════════════════════════════════════════════
# UTILITIES
# ═══════════════════════════════════════════════════════════════════════════

def make_label(text, size=19, bold=False, color=TEXT_PRI):
    lbl = QtWidgets.QLabel(text)
    weight = "600" if bold else "400"
    lbl.setStyleSheet(f"font-size:{size}px; font-weight:{weight}; color:{color};")
    return lbl


def make_section_label(text):
    lbl = QtWidgets.QLabel(text.upper())
    lbl.setStyleSheet(f"font-size:16px; font-weight:600; color:{TEXT_HINT}; letter-spacing:0.5px;")
    return lbl


def hline():
    line = QtWidgets.QFrame()
    line.setFrameShape(QtWidgets.QFrame.HLine)
    line.setStyleSheet(f"color:{GRAY_LINE}; border:none; border-top:1px solid {GRAY_LINE};")
    return line


def refresh_style(widget):
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.update()


def _fmt_num(v):
    """Format a number compactly for axis labels: 1234567 -> '1.23M', 12345 -> '12.3K'."""
    if v >= 1_000_000:
        return f"{v/1_000_000:.2f}M"
    if v >= 1_000:
        return f"{v/1_000:.1f}K"
    if isinstance(v, float):
        return f"{v:.1f}"
    return str(int(v))

def _spin(min_v, max_v, default, step=1, decimals=0):
    if decimals > 0:
        w = QtWidgets.QDoubleSpinBox()
        w.setDecimals(decimals)
        w.setSingleStep(step)
    else:
        w = QtWidgets.QSpinBox()
    w.setRange(min_v, max_v)
    w.setValue(default)
    return w


def _field(label_text, widget, tooltip=""):
    row = QtWidgets.QWidget()
    layout = QtWidgets.QVBoxLayout(row)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(3)
    lbl = make_label(label_text, size=16, color=TEXT_SEC)
    if tooltip:
        lbl.setToolTip(tooltip)
        widget.setToolTip(tooltip)
    layout.addWidget(lbl)
    layout.addWidget(widget)
    return row


def _grid(*fields, cols=2):
    container = QtWidgets.QWidget()
    grid = QtWidgets.QGridLayout(container)
    grid.setContentsMargins(0, 0, 0, 0)
    grid.setSpacing(12)
    for i, f in enumerate(fields):
        grid.addWidget(f, i // cols, i % cols)
    return container



_IUPAC: Dict[str, set] = {
    'A': {'A'}, 'C': {'C'}, 'G': {'G'}, 'T': {'T'},
    'R': {'A', 'G'}, 'Y': {'C', 'T'}, 'S': {'G', 'C'},
    'W': {'A', 'T'}, 'K': {'G', 'T'}, 'M': {'A', 'C'},
    'B': {'C', 'G', 'T'}, 'D': {'A', 'G', 'T'},
    'H': {'A', 'C', 'T'}, 'V': {'A', 'C', 'G'},
    'N': {'A', 'C', 'G', 'T'},
}

# IUPAC equalities list for edlib (same as original code)
_EDLIB_AMBIGUITY = [
    ("R", "A"), ("R", "G"), ("M", "A"), ("M", "C"),
    ("S", "C"), ("S", "G"), ("Y", "C"), ("Y", "T"),
    ("K", "G"), ("K", "T"), ("W", "A"), ("W", "T"),
    ("V", "A"), ("V", "C"), ("V", "G"),
    ("H", "A"), ("H", "C"), ("H", "T"),
    ("D", "A"), ("D", "G"), ("D", "T"),
    ("B", "C"), ("B", "G"), ("B", "T"),
    ("N", "A"), ("N", "G"), ("N", "C"), ("N", "T"),
]

_REVCOMP_TABLE = str.maketrans("ACGTRYSWKMBDHVNacgtryswkmbdhvn",
                                "TGCAYRSWMKVHDBNtgcayrswmkvhdbn")


# ---------------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------------

def _revcomp(seq: str) -> str:
    """Reverse complement respecting IUPAC ambiguities."""
    return seq.translate(_REVCOMP_TABLE)[::-1]


def _split_extract_cfg(cfg) -> Tuple[List[str], List[str], dict]:
    """
    Normalize an extraction config into (id_patterns, regex_patterns, normalize).

    ``cfg`` may be:
      • None  → legacy default (positional substring "_all.fa", no regex, no normalize)
      • a dict with keys "id_patterns" (list), "regex_patterns" (list),
        "normalize" (dict)
      • a plain str → treated as a single positional/substring pattern
    """
    if cfg is None:
        return ["_all.fa"], [], {}
    if isinstance(cfg, str):
        return ([cfg] if cfg else []), [], {}
    return (
        list(cfg.get("id_patterns") or []),
        list(cfg.get("regex_patterns") or []),
        dict(cfg.get("normalize") or {}),
    )


def _extract_sample_id(raw: str, id_patterns: List[str],
                       regex_patterns: List[str]) -> str:
    """
    Try each candidate pattern in order until one yields an ID.

    Order: regex candidates first (group 1 if present, else full match), then
    positional/substring candidates. A pattern that does not match is skipped so
    the next candidate gets a chance; only when none match do we fall back to
    the text before the first ';'.

    Positional pattern forms:
      • "pos:DELIM:N"    → text before the Nth occurrence of the literal DELIM
      • "posany:CHARS:N" → text before the Nth occurrence of any of the
                           single characters in CHARS.
    """
    import re as _re
    for rx in regex_patterns:
        if not rx:
            continue
        try:
            m = _re.search(rx, raw)
        except _re.error:
            continue
        if m:
            # Prefer group 1, but only if it actually participated in the match
            # (an optional group can leave group(1) == None while a later group
            # matched). Skip empty results so the next candidate gets a chance.
            cand = m.group(1) if (m.lastindex and m.group(1) is not None) else m.group(0)
            if cand:
                return cand

    for pat in id_patterns:
        if not pat:
            continue
        if pat.startswith("posany:"):
            try:
                # The char set itself may contain ':' (custom delimiter), so peel the
                # occurrence index off the RIGHT and keep everything else as the set.
                chars, _, n_str = pat[len("posany:"):].rpartition(":")
                n = int(n_str)
                if chars:
                    positions = [i for i, c in enumerate(raw) if c in chars]
                    if len(positions) >= n:
                        return raw[:positions[n - 1]]
            except Exception:
                continue
        elif pat.startswith("pos:"):
            try:
                # The separator may contain ':' too — split the index off the right.
                sep, _, n_str = pat[len("pos:"):].rpartition(":")
                n = int(n_str)
                if sep and sep in raw:
                    return sep.join(raw.split(sep)[:n])
            except Exception:
                continue
        elif pat in raw:
            return raw.split(pat)[0]

    return raw.split(";")[0]


def _normalize_id(sample_id: str, normalize: dict) -> str:
    """
    Normalize an extracted ID so equivalent IDs collide across files.

    Supported keys in ``normalize``:
      • "strip_regex" (str): substrings matching this regex are removed
      • "lowercase" (bool): casefold the ID
      • "strip_zeros" (bool): drop leading zeros inside numeric runs
    """
    if not normalize:
        return sample_id
    import re as _re
    sid = sample_id
    strip_rx = normalize.get("strip_regex") or ""
    if strip_rx:
        try:
            sid = _re.sub(strip_rx, "", sid)
        except _re.error:
            pass
    if normalize.get("lowercase"):
        sid = sid.lower()
    if normalize.get("strip_zeros"):
        sid = _re.sub(r"(?<!\d)0+(\d)", r"\1", sid)
    return sid.strip()


def _parse_fasta_header(header_line: str, cfg=None) -> Tuple[str, int, int, int, int]:
    """
    Extract (sample_id, length, coverage, ambs, gaps) from a FASTA header.

    ``cfg`` is an extraction config (see _split_extract_cfg). The ID is resolved
    by trying every candidate pattern in order and then normalized.
    For backward compatibility ``cfg`` may also be a plain pattern string or None.
    """
    raw = header_line.strip().lstrip(">")
    id_patterns, regex_patterns, normalize = _split_extract_cfg(cfg)
    sample_id = _extract_sample_id(raw, id_patterns, regex_patterns)
    sample_id = _normalize_id(sample_id, normalize)

    parts = raw.split(";")
    try:
        length = int(parts[1]) if len(parts) > 1 else 0
    except ValueError:
        length = 0
    try:
        coverage = int(parts[2]) if len(parts) > 2 else 0
    except ValueError:
        coverage = 0
    ambs = gaps = 0
    for p in parts[3:]:
        if p.startswith("ambs="):
            try:
                ambs = int(p.split("=")[1])
            except Exception:
                pass
        elif p.startswith("estgaps="):
            try:
                gaps = int(p.split("=")[1])
            except Exception:
                pass
    return sample_id, length, coverage, ambs, gaps


def _make_unique_labels(file_list: List[str]) -> List[str]:
    """
    Unique labels for each route. Use only the file name when not
    there is a collision; adds the parent directory when there is one.
    """
    raw = [os.path.basename(f) for f in file_list]
    raw_counts = Counter(raw)

    step1 = []
    for f, bn in zip(file_list, raw):
        if raw_counts[bn] > 1:
            parent = os.path.basename(os.path.dirname(os.path.abspath(f)))
            step1.append(f"{parent}/{bn}" if parent else bn)
        else:
            step1.append(bn)

    step1_counts = Counter(step1)
    if max(step1_counts.values(), default=1) == 1:
        return step1

    seen: Dict[str, int] = {}
    final = []
    for lbl in step1:
        if step1_counts[lbl] > 1:
            seen[lbl] = seen.get(lbl, 0) + 1
            final.append(f"{lbl} ({seen[lbl]})")
        else:
            final.append(lbl)
    return final


def _parse_fasta_file(path: str, cfg=None) -> Dict[str, Tuple[str, int, int, int, int]]:
    """
    Reads a FASTA file and returns dict[sample_id] = (seq, length, cov, ambs, gaps).
    ``cfg`` is the extraction config (see _split_extract_cfg / _parse_fasta_header).
    If there are duplicates in the same file, keep the one with the best quality
    (less both → fewer gaps → greater coverage).
    """
    result: Dict[str, Tuple[str, int, int, int, int]] = {}
    try:
        with open(path, encoding="utf-8", errors="replace") as fh:
            lines = fh.readlines()
    except Exception:
        return result

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith(">"):
            sid, length, cov, ambs, gaps = _parse_fasta_header(line, cfg)
            seq = ""
            j = i + 1
            while j < len(lines) and not lines[j].strip().startswith(">"):
                seq += lines[j].strip().upper()
                j += 1
            cand = (seq, length, cov, ambs, gaps)
            if sid not in result:
                result[sid] = cand
            else:
                old = result[sid]
                if (ambs, gaps, -cov) < (old[3], old[4], -old[2]):
                    result[sid] = cand
            i = j
        else:
            i += 1
    return result


# ---------------------------------------------------------------------------
# Sequence comparison (core)
# ---------------------------------------------------------------------------

def _align_pair(seq1: str, seq2: str):
    """
    Compare seq1 vs seq2 using edlib in NW (global alignment) mode.
    Retorna (d_noamb, d_amb):
      d_noamb – unambiguous edit distance
      d_amb – edit distance with IUPAC ambiguities
    NW ensures that the difference in length is reflected in the distance,
    which is correct for both Coding (same length) and non-Coding (variable length).
    """
    d_noamb = edlib.align(seq1, seq2, mode='NW', task='path')['editDistance']
    d_amb   = edlib.align(seq1, seq2, mode='NW', task='path',
                          additionalEqualities=_EDLIB_AMBIGUITY)['editDistance']
    return d_noamb, d_amb


def _compare_sequences(seq1: str, seq2: str) -> Tuple[str, int, int, bool]:
    """
    Compares two sequences and returns (state, d_amb, d_noamb, rc_used).

    state:
      'identical'  – identical character for character (d_noamb=0, d_amb=0)
      'compatible' – IUPAC-compatible (d_amb=0, d_noamb>0 = ambiguity positions)
      'different'  – incompatible (d_amb>0)
    d_amb:   edit distance with IUPAC equivalences (0 if identical/compatible)
    d_noamb: unambiguous edit distance (for compatible = IUPAC position count)
    rc_used: True if the match was found with the reverse complement
    """
    d_noamb, d_amb = _align_pair(seq1, seq2)
    if d_amb == 0:
        estado = 'identical' if d_noamb == 0 else 'compatible'
        return estado, 0, d_noamb, False
    rc1 = _revcomp(seq1)
    rc_noamb, rc_amb = _align_pair(rc1, seq2)
    if rc_amb == 0:
        estado = 'identical' if rc_noamb == 0 else 'compatible'
        return estado, 0, rc_noamb, True
    use_rc = rc_amb < d_amb
    best_amb   = rc_amb   if use_rc else d_amb
    best_noamb = rc_noamb if use_rc else d_noamb
    return 'different', best_amb, best_noamb, use_rc


def _iupac_compatible_simple(seq1: str, seq2: str) -> bool:
    """IUPAC position-to-position compatibility (same length only)."""
    if len(seq1) != len(seq2):
        return False
    for b1, b2 in zip(seq1, seq2):
        if not _IUPAC.get(b1, {b1}).intersection(_IUPAC.get(b2, {b2})):
            return False
    return True


def _hamming(seq1: str, seq2: str) -> Optional[int]:
    """Hamming distance; None if different lengths."""
    if len(seq1) != len(seq2):
        return None
    return sum(a != b for a, b in zip(seq1, seq2))


# ---------------------------------------------------------------------------
# Selection of the best barcode
# ---------------------------------------------------------------------------

# Entry: list of (basename, seq, length, cov, ambs, gaps)
_SeqEntry = Tuple[str, str, int, int, int, int]


def _best_barcode(entries: List[_SeqEntry]) -> Optional[_SeqEntry]:
    """
    Choose the entry with: less both → less gaps → greater coverage.
    Returns None if there is an exact tie between two or more entries.
    """
    if not entries:
        return None
    best_key = min((e[4], e[5], -e[3]) for e in entries)
    best = [e for e in entries if (e[4], e[5], -e[3]) == best_key]
    return best[0] if len(best) == 1 else None


# ---------------------------------------------------------------------------
# Writing outputs
# ---------------------------------------------------------------------------

_COLOR_MAP_HEX = {
    "Identical":           "#EAF3DE",
    "Compatible (IUPAC)": "#E6F1FB",
    "Different":          "#FCEBEB",
    "Only in reference": "#FAEEDA",
    "No reference":     "#F5F5F3",
}


def _write_fasta(path: str, entries: List[Tuple[str, str]]) -> None:
    """Write pairs (header, seq) to a FASTA file."""
    try:
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            for hdr, seq in entries:
                if seq:
                    f.write(f">{hdr}\n{seq}\n")
    except Exception:
        pass


def _write_outputs(
    rows: List[dict],
    headers: List[str],
    outdir: str,
    all_bns: List[str],
    ref_bn: Optional[str],
    seqs_store: Dict[str, Dict[str, Tuple[str, int, int, int, int]]],
) -> None:
    """
    Generate all output files:
      • summary.xlsx (with colors)
      • best_barcodes.fa
      • identical.fa
      • compatible_iupac.fa
      • different.fa
      • only_in_reference.fa
      • without_reference.fa
      • unique_<basename>.fa (one per file)
    """
    os.makedirs(outdir, exist_ok=True)

    # ── XLSX ────────────────────────────────────────────────────────────────

    xlsx_path = os.path.join(outdir, "summary.xlsx")
    try:
        wb = xlsxwriter.Workbook(xlsx_path)
        ws = wb.add_worksheet("Comparison")

        fmt_hdr = wb.add_format({
            "bold": True, "font_color": "#FFFFFF", "bg_color": "#185FA5",
            "border": 1, "border_color": "#2E6FAA",
            "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        fmt_by_estado = {}
        for estado, bg in _COLOR_MAP_HEX.items():
            fmt_by_estado[estado] = wb.add_format({
                "bg_color": bg, "border": 1,
                "border_color": "#D0CEC8", "valign": "vcenter",
            })
        fmt_unico = wb.add_format({
            "bg_color": "#EBEBEB", "border": 1,
            "border_color": "#D0CEC8", "valign": "vcenter",
        })
        fmt_default = wb.add_format({
            "border": 1, "border_color": "#D0CEC8", "valign": "vcenter",
        })
        fmt_id_bold = {
            estado: wb.add_format({
                "bold": True, "bg_color": bg, "border": 1,
                "border_color": "#D0CEC8", "valign": "vcenter",
            }) for estado, bg in _COLOR_MAP_HEX.items()
        }
        fmt_id_unico = wb.add_format({
            "bold": True, "bg_color": "#EBEBEB", "border": 1,
            "border_color": "#D0CEC8", "valign": "vcenter",
        })

        ws.set_row(0, 28)
        id_col = headers.index("ID") if "ID" in headers else -1
        for j, h in enumerate(headers):
            ws.write(0, j, h, fmt_hdr)

        for i, row in enumerate(rows, 1):
            ws.set_row(i, 18)
            estado = row.get("State", "")
            if estado in fmt_by_estado:
                row_fmt    = fmt_by_estado[estado]
                row_id_fmt = fmt_id_bold.get(estado, row_fmt)
            elif estado.startswith("Unique"):
                row_fmt    = fmt_unico
                row_id_fmt = fmt_id_unico
            else:
                row_fmt    = fmt_default
                row_id_fmt = fmt_default

            for j, h in enumerate(headers):
                val = row.get(h, "")
                fmt = row_id_fmt if j == id_col else row_fmt
                ws.write(i, j, val, fmt)

        for j, h in enumerate(headers):
            col_vals = [str(row.get(h, "") or "") for row in rows]
            max_len  = max((len(v) for v in col_vals), default=4)
            max_len  = max(max_len, len(h))
            ws.set_column(j, j, min(max_len + 2, 52))

        ws.freeze_panes(1, 1)
        wb.close()
    except Exception as exc:
        print(f"[XLSX] Writing error: {exc}")

    # ── Helpers ─────────────────────────────────────────────────────────────
    def get_seq(bn: str, sid: str):
        entry = seqs_store.get(bn, {}).get(sid)
        return (entry[0], entry[3], entry[2]) if entry else None

    # ── Sort rows and accumulate FASTA entries ───────────────────────────
    best_entries:      List[Tuple[str, str]] = []
    identical_entries: List[Tuple[str, str]] = []
    compat_entries:    List[Tuple[str, str]] = []
    diff_entries:      List[Tuple[str, str]] = []
    only_ref_entries:  List[Tuple[str, str]] = []
    no_ref_entries:    List[Tuple[str, str]] = []
    unique_entries:    Dict[str, List[Tuple[str, str]]] = {bn: [] for bn in all_bns}

    for row in rows:
        sid    = row["ID"]
        estado = row.get("State", "")
        best_bn = row.get("Best_run", "")

        # Best overall barcode
        candidates: List[_SeqEntry] = []
        for bn in all_bns:
            entry = seqs_store.get(bn, {}).get(sid)
            if entry:
                seq, length, cov, ambs, gaps = entry
                candidates.append((bn, seq, length, cov, ambs, gaps))
        if candidates:
            if estado == "Different" and best_bn and best_bn != "Tie":
                chosen = next((c for c in candidates if c[0] == best_bn),
                              candidates[0])
            else:
                chosen = min(candidates, key=lambda c: (c[4], c[5], -c[3]))
            best_entries.append((f"{sid};best_from={chosen[0]}", chosen[1]))

        # Files by category
        if estado == "Identical":
            for bn in all_bns:
                r = get_seq(bn, sid)
                if r:
                    identical_entries.append((f"{sid};src={bn}", r[0]))
                    break
        elif estado == "Compatible (IUPAC)":
            for bn in all_bns:
                r = get_seq(bn, sid)
                if r:
                    compat_entries.append((f"{sid};src={bn}", r[0]))
                    break
        elif estado == "Different":
            for bn in all_bns:
                r = get_seq(bn, sid)
                if r:
                    diff_entries.append((f"{sid};src={bn}", r[0]))
        elif estado == "Only in reference":
            if ref_bn:
                r = get_seq(ref_bn, sid)
                if r:
                    only_ref_entries.append((sid, r[0]))
        elif estado == "No reference":
            for bn in all_bns:
                if bn == ref_bn:
                    continue
                r = get_seq(bn, sid)
                if r:
                    no_ref_entries.append((f"{sid};src={bn}", r[0]))
                    break
        elif estado.startswith("Unique in "):
            src_bn = estado.replace("Unique in ", "")
            r = get_seq(src_bn, sid)
            if r and src_bn in unique_entries:
                unique_entries[src_bn].append((sid, r[0]))

    # ── Write FASTAs ─────────────────────────── ───────────────────────────
    _write_fasta(os.path.join(outdir, "best_barcodes.fa"), best_entries)
    if identical_entries:
        _write_fasta(os.path.join(outdir, "identical.fa"), identical_entries)
    if compat_entries:
        _write_fasta(os.path.join(outdir, "compatible_iupac.fa"), compat_entries)
    if diff_entries:
        _write_fasta(os.path.join(outdir, "different.fa"), diff_entries)
    if only_ref_entries:
        _write_fasta(os.path.join(outdir, "only_in_reference.fa"), only_ref_entries)
    if no_ref_entries:
        _write_fasta(os.path.join(outdir, "without_reference.fa"), no_ref_entries)
    for bn, entries in unique_entries.items():
        if entries:
            safe = bn.replace("/", "_").replace("\\", "_")
            _write_fasta(os.path.join(outdir, f"unique_{safe}.fa"), entries)


# ---------------------------------------------------------------------------
# Worker: everyone vs everyone
# ---------------------------------------------------------------------------


# Color del marco/separador de las gráficas (solo afecta a estos gráficos,
# independiente de GRAY_LINE del resto de la interfaz). Cámbialo aquí.
CHART_FRAME_COLOR = "#B7BABA"


class _HistWidget(QtWidgets.QWidget):
    """Histogram drawn with QPainter.
    Bins are pre-computed in set_data() so paintEvent is O(n_bins), not O(n_values).
    render_to_pixmap() draws at any resolution with scaled fonts for crisp PDF export.
    """

    def __init__(self, title="", x_label="", color=BLUE_MID, n_bins=60, parent=None):
        super().__init__(parent)
        self._title   = title
        self._x_label = x_label
        self._color   = QtGui.QColor(color)
        self._n_bins  = n_bins
        self._counts: List[int] = []
        self._v_min   = 0.0
        self._v_max   = 1.0
        self._max_cnt = 1
        self.setMinimumHeight(210)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

    def set_data(self, values, display_min=None, display_max=None):
        """Pre-compute histogram bins. display_min/max clip the x-axis without
        affecting statistics — only the visible range of the chart changes."""
        if not values:
            self._counts = []
            self.update()
            return

        v_min = display_min if display_min is not None else min(values)
        v_max = display_max if display_max is not None else max(values)
        if v_min >= v_max:
            v_max = v_min + 1

        self._v_min = v_min
        self._v_max = v_max
        span = v_max - v_min

        counts = [0] * self._n_bins
        for v in values:
            if v < v_min or v > v_max:
                continue
            idx = int((v - v_min) / span * self._n_bins)
            if idx >= self._n_bins:
                idx = self._n_bins - 1
            counts[idx] += 1

        self._counts  = counts
        self._max_cnt = max(counts) or 1
        self.update()

    def set_precomputed(self, counts: list, v_min: float, v_max: float):
        """Accept bins already computed in the worker thread (avoids O(n) on main thread)."""
        if not counts:
            self._counts = []
            self.update()
            return
        self._v_min   = v_min
        self._v_max   = v_max if v_max > v_min else v_min + 1
        self._counts  = counts
        self._max_cnt = max(counts) or 1
        self.update()

    def render_to_pixmap(self, w: int, h: int, scale: float = 1.0) -> QtGui.QPixmap:
        pix = QtGui.QPixmap(w, h)
        pix.fill(QtCore.Qt.white)
        p = QtGui.QPainter(pix)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        p.setRenderHint(QtGui.QPainter.TextAntialiasing)
        self._draw(p, w, h, scale)
        p.end()
        return pix

    def paintEvent(self, event):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        self._draw(p, self.width(), self.height(), 1.0)
        p.end()

    def _draw(self, p: QtGui.QPainter, W: int, H: int, s: float = 1.0):
        PL = int(58 * s); PR = int(14 * s); PT = int(28 * s); PB = int(44 * s)

        p.fillRect(0, 0, W, H, QtGui.QColor("#FFFFFF"))

        font_t = QtGui.QFont()
        font_t.setBold(True)
        font_t.setPointSizeF(9 * s)
        p.setFont(font_t)
        p.setPen(QtGui.QColor(TEXT_PRI))
        p.drawText(QtCore.QRect(PL, 4, W - PL - PR, PT - 4),
                   QtCore.Qt.AlignCenter, self._title)

        cw, ch = W - PL - PR, H - PT - PB
        p.setPen(QtGui.QPen(QtGui.QColor(CHART_FRAME_COLOR), max(1, int(s))))
        p.drawRect(PL, PT, cw, ch)

        if not self._counts:
            font_e = QtGui.QFont()
            font_e.setPointSizeF(8 * s)
            p.setFont(font_e)
            p.setPen(QtGui.QColor(TEXT_HINT))
            p.drawText(QtCore.QRect(PL, PT, cw, ch), QtCore.Qt.AlignCenter, "No data")
            return

        gap   = max(1, int(s))
        bar_w = cw / self._n_bins
        fill  = QtGui.QColor(self._color)
        fill.setAlpha(190)
        p.setPen(QtCore.Qt.NoPen)
        p.setBrush(fill)
        for i, cnt in enumerate(self._counts):
            bh = int(cnt / self._max_cnt * ch)
            bx = PL + int(i * bar_w)
            p.drawRect(bx + gap, PT + ch - bh, max(1, int(bar_w) - gap), bh)

        span = self._v_max - self._v_min
        lw   = int(52 * s); lh = int(14 * s)
        font_s = QtGui.QFont()
        font_s.setPointSizeF(7 * s)
        p.setFont(font_s)
        p.setPen(QtGui.QColor(TEXT_SEC))
        for i in range(5):
            xv = self._v_min + span * i / 4
            px = PL + int(cw * i / 4)
            p.drawText(QtCore.QRect(px - lw // 2, PT + ch + int(4 * s), lw, lh),
                       QtCore.Qt.AlignCenter, _fmt_num(xv))
        for i in range(4):
            yv = int(self._max_cnt * i / 3)
            py = PT + ch - int(ch * i / 3)
            p.drawText(QtCore.QRect(1, py - lh // 2, PL - int(4 * s), lh),
                       QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter, _fmt_num(yv))

        if self._x_label:
            font_xl = QtGui.QFont()
            font_xl.setPointSizeF(8 * s)
            p.setFont(font_xl)
            p.setPen(QtGui.QColor(TEXT_SEC))
            p.drawText(QtCore.QRect(PL, H - int(16 * s), cw, int(16 * s)),
                       QtCore.Qt.AlignCenter, self._x_label)


class _ScatterWidget(QtWidgets.QWidget):
    """Scatter plot drawn with QPainter.
    Points are pre-filtered to the display range in set_data() so paintEvent
    only iterates the visible subset.
    render_to_pixmap() draws at any resolution with scaled fonts for crisp PDF export.
    """

    def __init__(self, title="", x_label="", y_label="", color=BLUE_MID, parent=None):
        super().__init__(parent)
        self._title   = title
        self._x_label = x_label
        self._y_label = y_label
        self._color   = QtGui.QColor(color)
        self._pts: List[Tuple[float, float]] = []
        self._x_min = 0.0
        self._x_max = 1.0
        self._y_min = 0.0
        self._y_max = 1.0
        self.setMinimumHeight(210)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

    def set_data(self, xs, ys, x_min=None, x_max=None, y_min=None, y_max=None):
        """Pre-filter points to the display range. x_min/x_max clip the x-axis
        without affecting statistics — only the visible range of the chart changes.
        y_min/y_max let the caller pin the Y axis (e.g. the full-data Q range) so the
        scatter lines up with its companion histogram instead of deriving the range
        from the sampled subset."""
        if not xs:
            self._pts = []
            self.update()
            return

        xlo = x_min if x_min is not None else min(xs)
        xhi = x_max if x_max is not None else max(xs)
        if xlo >= xhi:
            xhi = xlo + 1

        pts = [(x, y) for x, y in zip(xs, ys) if xlo <= x <= xhi]

        self._x_min = xlo
        self._x_max = xhi
        self._y_min = y_min if y_min is not None else min((pt[1] for pt in pts), default=0.0)
        self._y_max = y_max if y_max is not None else max((pt[1] for pt in pts), default=1.0)
        if self._y_min >= self._y_max:
            self._y_max = self._y_min + 1
        self._pts = pts
        self.update()

    def render_to_pixmap(self, w: int, h: int, scale: float = 1.0) -> QtGui.QPixmap:
        pix = QtGui.QPixmap(w, h)
        pix.fill(QtCore.Qt.white)
        p = QtGui.QPainter(pix)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        p.setRenderHint(QtGui.QPainter.TextAntialiasing)
        self._draw(p, w, h, scale)
        p.end()
        return pix

    def paintEvent(self, event):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        self._draw(p, self.width(), self.height(), 1.0)
        p.end()

    def _draw(self, p: QtGui.QPainter, W: int, H: int, s: float = 1.0):
        PL = int(58 * s); PR = int(14 * s); PT = int(28 * s); PB = int(44 * s)

        p.fillRect(0, 0, W, H, QtGui.QColor("#FFFFFF"))

        font_t = QtGui.QFont()
        font_t.setBold(True)
        font_t.setPointSizeF(9 * s)
        p.setFont(font_t)
        p.setPen(QtGui.QColor(TEXT_PRI))
        p.drawText(QtCore.QRect(PL, 4, W - PL - PR, PT - 4),
                   QtCore.Qt.AlignCenter, self._title)

        cw, ch = W - PL - PR, H - PT - PB
        p.setPen(QtGui.QPen(QtGui.QColor(CHART_FRAME_COLOR), max(1, int(s))))
        p.drawRect(PL, PT, cw, ch)

        if not self._pts:
            font_e = QtGui.QFont()
            font_e.setPointSizeF(8 * s)
            p.setFont(font_e)
            p.setPen(QtGui.QColor(TEXT_HINT))
            p.drawText(QtCore.QRect(PL, PT, cw, ch), QtCore.Qt.AlignCenter, "No data")
            return

        xr = self._x_max - self._x_min
        yr = self._y_max - self._y_min

        dot = QtGui.QColor(self._color)
        dot.setAlpha(60)
        p.setPen(QtCore.Qt.NoPen)
        p.setBrush(dot)
        r = max(2, int(2 * s))
        for x, y in self._pts:
            px = PL + int((x - self._x_min) / xr * cw)
            py = PT + ch - int((y - self._y_min) / yr * ch)
            p.drawEllipse(QtCore.QPoint(px, py), r, r)

        lw   = int(52 * s); lh = int(14 * s)
        font_s = QtGui.QFont()
        font_s.setPointSizeF(7 * s)
        p.setFont(font_s)
        p.setPen(QtGui.QColor(TEXT_SEC))
        for i in range(5):
            xv = self._x_min + xr * i / 4
            px = PL + int(cw * i / 4)
            p.drawText(QtCore.QRect(px - lw // 2, PT + ch + int(4 * s), lw, lh),
                       QtCore.Qt.AlignCenter, _fmt_num(xv))
        for i in range(4):
            yv = self._y_min + yr * i / 3
            py = PT + ch - int(ch * i / 3)
            p.drawText(QtCore.QRect(1, py - lh // 2, PL - int(4 * s), lh),
                       QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter, _fmt_num(yv))

        if self._x_label:
            font_xl = QtGui.QFont()
            font_xl.setPointSizeF(8 * s)
            p.setFont(font_xl)
            p.setPen(QtGui.QColor(TEXT_SEC))
            p.drawText(QtCore.QRect(PL, H - int(16 * s), cw, int(16 * s)),
                       QtCore.Qt.AlignCenter, self._x_label)

        if self._y_label:
            p.save()
            p.translate(int(11 * s), PT + ch // 2)
            p.rotate(-90)
            font_yl = QtGui.QFont()
            font_yl.setPointSizeF(8 * s)
            p.setFont(font_yl)
            p.setPen(QtGui.QColor(TEXT_SEC))
            p.drawText(QtCore.QRect(-ch // 2, -lh // 2, ch, lh),
                       QtCore.Qt.AlignCenter, self._y_label)
            p.restore()


# ═══════════════════════════════════════════════════════════════════════════
# FASTQ INSPECTOR – module-level helpers for multiprocessing
# (must be at module level so they are picklable on Windows)
# ═══════════════════════════════════════════════════════════════════════════

def _fq_sync_record(fh_bin) -> bool:
    """Advance a binary file handle to the start of the next valid FASTQ record.
    Needed when seeking to an arbitrary byte offset inside the file.
    Returns True if a record was found, False on EOF."""
    for _ in range(16):
        line = fh_bin.readline()
        if not line:
            return False
        if not line.startswith(b'@'):
            continue
        # Verify: next line is seq, then '+', then qual of matching length
        pos   = fh_bin.tell()
        seq   = fh_bin.readline()
        plus  = fh_bin.readline()
        qual  = fh_bin.readline()
        if plus.startswith(b'+') and len(seq.strip()) == len(qual.strip()):
            fh_bin.seek(pos - len(line))   # rewind to the '@' header
            return True
        # False positive (@ in a quality line) — rewind past the 3 look-ahead
        # lines so the next scan re-examines them; otherwise a real header
        # consumed during verification would be skipped.
        fh_bin.seek(pos)
    return False


def _fq_chunk_task(path: str, start: int, end: int) -> dict:
    """Process reads in byte range [start, end) of an UNCOMPRESSED FASTQ file.
    Returns a partial-stats dict that _FastqInspectorWorker aggregates afterward.
    Defined at module level so ProcessPoolExecutor can pickle it on Windows."""
    lengths:  List[int]   = []
    q_scores: List[float] = []
    gc_pcts:  List[float] = []
    total_bases = 0
    q_sum = gc_sum = 0.0
    len_lt500 = len_500_2k = len_gt2k = 0
    q_cnt_10 = q_cnt_15 = q_cnt_20 = 0

    _EMPTY = {"n_reads": 0, "lengths": [], "q_scores": [], "gc_pcts": [],
              "total_bases": 0, "q_sum": 0.0, "gc_sum": 0.0,
              "len_lt500": 0, "len_500_2k": 0, "len_gt2k": 0,
              "q_cnt_10": 0, "q_cnt_15": 0, "q_cnt_20": 0}

    try:
        with open(path, "rb", buffering=1 << 20) as fh_bin:
            if start > 0:
                fh_bin.seek(start)
                if not _fq_sync_record(fh_bin):
                    return _EMPTY

            # Partition records by their START offset: this worker owns every
            # record beginning in [start, end). Using readline() + an explicit
            # record-start offset (instead of `for line in fh` + tell(), whose
            # read-ahead buffer makes tell() inaccurate) avoids both dropping the
            # record that straddles the boundary and double-counting it — the next
            # worker resyncs to the first record starting at/after `end`.
            while True:
                rec_start = fh_bin.tell()
                if end > 0 and rec_start >= end:
                    break
                raw_header = fh_bin.readline()
                if not raw_header:
                    break
                if not raw_header.startswith(b'@'):
                    continue                          # skip malformed / mid-record
                raw_seq  = fh_bin.readline()
                fh_bin.readline()                    # '+'
                raw_qual = fh_bin.readline()

                seq  = raw_seq.rstrip(b'\r\n')
                qual = raw_qual.rstrip(b'\r\n')
                L    = len(seq)
                ql   = len(qual) or 1
                if L == 0:
                    continue

                # Fast GC — str.count() is C-speed; avoids per-char Python loop
                gc_cnt = (seq.count(b'G') + seq.count(b'C') +
                          seq.count(b'g') + seq.count(b'c'))
                gc     = gc_cnt / L * 100.0

                # Fast Q-mean — bytes iterate as ints; no ord() call needed
                q_mean = (sum(qual) - 33 * ql) / ql

                lengths.append(L)
                q_scores.append(q_mean)
                gc_pcts.append(gc)

                total_bases += L
                q_sum       += q_mean
                gc_sum      += gc

                if L < 500:      len_lt500  += 1
                elif L <= 2000:  len_500_2k += 1
                else:            len_gt2k   += 1

                if q_mean >= 10: q_cnt_10 += 1
                if q_mean >= 15: q_cnt_15 += 1
                if q_mean >= 20: q_cnt_20 += 1

    except Exception:
        pass   # partial result still aggregated

    return {"n_reads": len(lengths), "lengths": lengths, "q_scores": q_scores,
            "gc_pcts": gc_pcts, "total_bases": total_bases,
            "q_sum": q_sum, "gc_sum": gc_sum,
            "len_lt500": len_lt500, "len_500_2k": len_500_2k, "len_gt2k": len_gt2k,
            "q_cnt_10": q_cnt_10, "q_cnt_15": q_cnt_15, "q_cnt_20": q_cnt_20}


# ═══════════════════════════════════════════════════════════════════════════
# FASTQ INSPECTOR – background worker
# ═══════════════════════════════════════════════════════════════════════════


class DropZone(QtWidgets.QFrame):
    fileDropped = QtCore.pyqtSignal(str)

    def __init__(self, label, hint="", extensions=None, parent=None):
        super().__init__(parent)
        self.setObjectName("drop_zone")
        self.setAcceptDrops(True)
        self._extensions = extensions or []
        self._filepath = ""
        self._original_label = label
        self._original_hint = hint

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(4)
        layout.setAlignment(QtCore.Qt.AlignCenter)

        self._label = make_label(label, color=TEXT_SEC)
        self._label.setAlignment(QtCore.Qt.AlignCenter)
        self._hint = make_label(hint, size=15, color=TEXT_HINT)
        self._hint.setAlignment(QtCore.Qt.AlignCenter)
        self._file_lbl = make_label("", size=15, bold=True, color=GREEN)
        self._file_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._file_lbl.hide()

        self._drag_icon_lbl = QtWidgets.QLabel()
        self._drag_icon_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._drag_icon_lbl.hide()

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.setSpacing(8)
        btn_row.setAlignment(QtCore.Qt.AlignCenter)

        self._browse_btn = QtWidgets.QPushButton("Browse...")
        self._browse_btn.setObjectName("secondary_btn")
        self._browse_btn.setFixedWidth(120)
        self._browse_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {BLUE};
                border: 1px solid {BLUE_MID};
                border-radius: 8px;
                padding: 7px 16px;
                font-size: 18px;
            }}
            QPushButton:hover {{
                background-color: {BLUE_LIGHT};
                color: {BLUE};
                border-color: {BLUE};
            }}
            QPushButton:pressed {{
                background-color: {BLUE_MID};
                color: white;
            }}
        """)
        self._browse_btn.clicked.connect(self._browse)

        self._clear_btn = QtWidgets.QPushButton("✕ Remove")
        self._clear_btn.setObjectName("danger_btn")
        self._clear_btn.setFixedWidth(120)
        self._clear_btn.clicked.connect(self.clear)
        self._clear_btn.hide()

        btn_row.addWidget(self._browse_btn)
        btn_row.addWidget(self._clear_btn)

        layout.addWidget(self._drag_icon_lbl)
        layout.addWidget(self._label)
        layout.addWidget(self._hint)
        layout.addWidget(self._file_lbl)
        layout.addSpacing(8)
        layout.addLayout(btn_row)

    def retranslateUi(self):
        ctx = "DropZone"
        self._browse_btn.setText(_tr(ctx, "Browse..."))
        self._clear_btn.setText(_tr(ctx, "✕ Remove"))
        self._hint.setText(_tr(ctx, self._original_hint))
        if not self._filepath:
            self._label.setText(_tr(ctx, self._original_label))
        else:
            self._label.setText(_tr(ctx, "File loaded"))

    def changeEvent(self, event):
        if event.type() == QtCore.QEvent.LanguageChange:
            self.retranslateUi()
        super().changeEvent(event)

    def _browse(self):
        ctx = "DropZone"
        ext_str = " ".join(f"*{e}" for e in self._extensions) if self._extensions else "*"
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, _tr(ctx, "Select file"), "", f"Files ({ext_str})"
        )
        if path:
            self._set_file(path)

    def _set_file(self, path):
        self._filepath = path
        name = os.path.basename(path)
        size = os.path.getsize(path)
        if size >= 1_073_741_824:
            size_str = f"{size/1_073_741_824:.2f} GB"
        elif size >= 1_048_576:
            size_str = f"{size/1_048_576:.1f} MB"
        elif size >= 1_024:
            size_str = f"{size/1_024:.1f} KB"
        else:
            size_str = f"{size} B"
        self._file_lbl.setText(f"{name}  ·  {size_str}")
        self._label.setText(_tr("DropZone", "File loaded"))
        self._file_lbl.show()
        self._clear_btn.show()
        self.setProperty("filled", "true")
        self.setCursor(QtCore.Qt.PointingHandCursor)
        refresh_style(self)
        self.fileDropped.emit(path)

    def clear(self):
        self._filepath = ""
        self._file_lbl.hide()
        self._clear_btn.hide()
        self._label.setText(_tr("DropZone", self._original_label))
        self.setProperty("filled", "false")
        refresh_style(self)
        self.fileDropped.emit("")

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self.setProperty("dragging", "true")
            urls = e.mimeData().urls()
            if urls:
                path = urls[0].toLocalFile()
                if path:
                    icon = QtWidgets.QFileIconProvider().icon(QtCore.QFileInfo(path))
                    self._drag_icon_lbl.setPixmap(icon.pixmap(48, 48))
                    self._drag_icon_lbl.show()
            refresh_style(self)

    def dragLeaveEvent(self, e):
        self.setProperty("dragging", "false")
        self._drag_icon_lbl.hide()
        refresh_style(self)

    def dropEvent(self, e):
        self.setProperty("dragging", "false")
        self._drag_icon_lbl.hide()
        refresh_style(self)
        for url in e.mimeData().urls():
            path = url.toLocalFile()
            if self._extensions:
                if any(path.lower().endswith(ext) for ext in self._extensions):
                    self._set_file(path)
                    return
            else:
                self._set_file(path)
                return

    @property
    def filepath(self):
        return self._filepath



class MultiDropZone(QtWidgets.QFrame):
    filesDropped = QtCore.pyqtSignal(list)

    _ROW_H   = 60   # px per file row (content + spacing)
    _CHROME  = 200   # label + buttons + layout margins/spacings (6+24+8+8+36+6)
    _EMPTY_H = 200   # height when no files

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("drop_zone")
        self.setAcceptDrops(True)
        self.setFixedHeight(self._EMPTY_H)
        self._files = []  # Lista acumulativa
        self._seq_cache: dict = {}  # path → seq count (evita reconteo al refrescar)
        self._supported_extensions = (".fa", ".fas", ".fasta", ".fa.gz", ".fas.gz", ".fasta.gz", ".fna", ".fna.gz")

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(10, 6, 10, 6)
        layout.setSpacing(8)

        # ── Main label (compact, with hint as tooltip) ──
        self._lbl = make_label(
            "Drag/Add FASTA files here",
            size=18, color=TEXT_SEC)
        self._lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._lbl.setToolTip("You can drag multiple files at once or one at a time.")

        # ── Hint (hidden, kept for compatibility) ──
        self._hint = make_label("", size=14, color=TEXT_HINT)
        self._hint.hide()

        self._drag_icon_lbl = QtWidgets.QLabel()
        self._drag_icon_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._drag_icon_lbl.hide()

        # ── Container for the list of files ──
        self._files_container = QtWidgets.QWidget()
        self._files_layout = QtWidgets.QVBoxLayout(self._files_container)
        self._files_layout.setContentsMargins(0, 0, 0, 0)
        self._files_layout.setSpacing(8)
        self._files_container.hide()

        # ── Buttons ──
        btn_layout = QtWidgets.QHBoxLayout()
        btn_layout.setSpacing(8)
        btn_layout.setAlignment(QtCore.Qt.AlignCenter)
        
        self._browse_btn = QtWidgets.QPushButton("Add")
        self._browse_btn.setObjectName("secondary_btn")
        self._browse_btn.setFixedWidth(130)
        self._browse_btn.clicked.connect(self._browse_files)
        self._lbl_src_empty = "Drag/Add FASTA files here"
        self._lbl_src_filled = "Uploaded"

        self._clear_btn = QtWidgets.QPushButton("Clear")
        self._clear_btn.setObjectName("danger_btn")
        self._clear_btn.setFixedWidth(140)
        self._clear_btn.clicked.connect(self.clear)
        self._clear_btn.hide()
        
        btn_layout.addWidget(self._browse_btn)
        btn_layout.addWidget(self._clear_btn)

        # ── Add everything to the main layout ──
        # Spacers colapsables: expanden en estado vacío para centrar, se colapsan cuando hay archivos
        _exp = QtWidgets.QSizePolicy
        self._top_spacer = QtWidgets.QSpacerItem(0, 0, _exp.Minimum, _exp.Expanding)
        self._bot_spacer = QtWidgets.QSpacerItem(0, 0, _exp.Minimum, _exp.Expanding)
        layout.addSpacerItem(self._top_spacer)
        layout.addWidget(self._drag_icon_lbl)
        layout.addWidget(self._lbl)
        layout.addWidget(self._files_container)
        layout.addLayout(btn_layout)
        layout.addSpacerItem(self._bot_spacer)

    def retranslateUi(self):
        ctx = "MultiDropZone"
        self._browse_btn.setText(_tr(ctx, "Add"))
        self._clear_btn.setText(_tr(ctx, "Clear"))
        if len(self._files) == 0:
            self._lbl.setText(_tr(ctx, self._lbl_src_empty))
        else:
            total_seqs = sum(self._seq_cache.get(f, 0) for f in self._files)
            self._lbl.setText(
                f"{_tr(ctx, self._lbl_src_filled)} ({len(self._files)} files · {total_seqs:,} seqs)"
            )

    def changeEvent(self, event):
        if event.type() == QtCore.QEvent.LanguageChange:
            self.retranslateUi()
        super().changeEvent(event)

    @staticmethod
    def _count_seqs(filepath):
        """Count FASTA sequences by counting lines starting with '>'."""
        try:
            opener = __import__("gzip").open if filepath.endswith(".gz") else open
            with opener(filepath, "rt", encoding="utf-8", errors="replace") as fh:
                return sum(1 for ln in fh if ln.startswith(">"))
        except Exception:
            return 0

    def _create_file_row(self, filepath, index):
        """Create a row with folder/file, sequence count and remove button."""
        parent_dir = os.path.basename(os.path.dirname(os.path.abspath(filepath)))
        display_name = f"{parent_dir}/{os.path.basename(filepath)}" if parent_dir else os.path.basename(filepath)

        row_widget = QtWidgets.QWidget()
        row_widget.setObjectName("file_row")
        row_widget.setStyleSheet(f"""
            QWidget#file_row {{
                background-color: {GRAY_BG};
                border-radius: 7px;
                border: 1px solid {GRAY_LINE};
            }}
            QWidget#file_row:hover {{
                background-color: {BLUE_LIGHT};
                border-color: #B8D4F0;
            }}
        """)
        row_widget.setToolTip(filepath)

        row_layout = QtWidgets.QHBoxLayout(row_widget)
        row_layout.setContentsMargins(12, 8, 10, 8)
        row_layout.setSpacing(10)

        # Icon
        icon_lbl = make_label("🧬", size=15)
        icon_lbl.setFixedWidth(24)

        # Center column: name + count
        center = QtWidgets.QVBoxLayout()
        center.setSpacing(1)
        name_lbl = make_label(display_name, size=15, color=TEXT_PRI)
        name_lbl.setWordWrap(False)

        n_seqs = self._seq_cache.get(filepath)
        if n_seqs is None:
            n_seqs = self._count_seqs(filepath)
            self._seq_cache[filepath] = n_seqs
        seq_lbl = make_label(
            f"{n_seqs:,} sequences",
            size=14, color=TEXT_HINT
        )
        center.addWidget(name_lbl)
        center.addWidget(seq_lbl)

        # remove button
        remove_btn = QtWidgets.QPushButton("✕")
        remove_btn.setFixedSize(26, 26)
        remove_btn.setToolTip("Remove file")
        remove_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {TEXT_HINT};
                border: none;
                border-radius: 5px;
                font-size: 13px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {RED_LT};
                color: {RED};
            }}
            QPushButton:pressed {{ background-color: {RED}; color: white; }}
        """)
        remove_btn.clicked.connect(lambda checked, idx=index: self.remove_file(idx))

        row_layout.addWidget(icon_lbl)
        row_layout.addLayout(center, 1)
        row_layout.addWidget(remove_btn)

        return row_widget

    def _adjust_height(self):
        n = len(self._files)
        if n == 0:
            self.setFixedHeight(self._EMPTY_H)
        else:
            content_h = n * self._ROW_H - (n - 1) * 8
            self.setFixedHeight(self._CHROME + content_h)

    def _update_display(self):
        """Update the file list in the interface"""
        n = len(self._files)

        if n == 0:
            self._files_container.hide()
            self._clear_btn.hide()
            self._lbl.setText(_tr("MultiDropZone", self._lbl_src_empty))
            self.setProperty("filled", "false")
        else:
            # Clear all rows
            while self._files_layout.count() > 0:
                item = self._files_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()

            # Add each file as a row
            for i, f in enumerate(self._files):
                row = self._create_file_row(f, i)
                self._files_layout.addWidget(row)

            self._files_container.show()
            self._clear_btn.show()
            total_seqs = sum(self._seq_cache.get(f, 0) for f in self._files)
            self._lbl.setText(
                f"{_tr('MultiDropZone', self._lbl_src_filled)} ({n} files · {total_seqs:,} seqs)"
            )
            self.setProperty("filled", "true")

        # Colapsar spacers cuando hay archivos; expandir para centrar cuando está vacío
        _sp = QtWidgets.QSizePolicy
        if n == 0:
            self._top_spacer.changeSize(0, 0, _sp.Minimum, _sp.Expanding)
            self._bot_spacer.changeSize(0, 0, _sp.Minimum, _sp.Expanding)
        else:
            self._top_spacer.changeSize(0, 0, _sp.Minimum, _sp.Fixed)
            self._bot_spacer.changeSize(0, 0, _sp.Minimum, _sp.Fixed)
        self.layout().invalidate()

        self._adjust_height()
        refresh_style(self)
        refresh_style(self._files_container)

    def _browse_files(self):
        """Open dialog to select multiple files"""
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, _tr("MultiDropZone", "Select FASTA files"), "",
            "FASTA files (*.fa *.fas *.fasta *.fa.gz *.fas.gz *.fasta.gz *.fna *.fna.gz);;All (*)"
        )
        if files:
            self._add_files(files)

    def _add_files(self, new_files):
        """Add new files without overwriting existing ones"""
        added = 0
        for f in new_files:
            if f not in self._files:
                self._files.append(f)
                added += 1
        
        if added > 0:
            self._update_display()
            self.filesDropped.emit(self._files.copy())

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self.setProperty("dragging", "true")
            urls = e.mimeData().urls()
            if urls:
                path = urls[0].toLocalFile()
                if path:
                    icon = QtWidgets.QFileIconProvider().icon(QtCore.QFileInfo(path))
                    self._drag_icon_lbl.setPixmap(icon.pixmap(48, 48))
                    self._drag_icon_lbl.show()
            refresh_style(self)

    def dragLeaveEvent(self, e):
        self.setProperty("dragging", "false")
        self._drag_icon_lbl.hide()
        refresh_style(self)

    def dropEvent(self, e):
        self.setProperty("dragging", "false")
        self._drag_icon_lbl.hide()
        refresh_style(self)
        paths = []
        for url in e.mimeData().urls():
            path = url.toLocalFile()
            if any(path.lower().endswith(ext) for ext in self._supported_extensions):
                paths.append(path)
        
        if paths:
            self._add_files(paths)

    def clear(self):
        self._files = []
        self._seq_cache = {}
        self._update_display()
        self.filesDropped.emit([])

    def remove_file(self, index):
        """Delete a specific file by index"""
        if 0 <= index < len(self._files):
            removed = self._files.pop(index)
            self._update_display()
            self.filesDropped.emit(self._files.copy())
            return removed
        return None

    @property
    def files(self):
        return self._files.copy()



class BlastPanel(QtWidgets.QWidget):
    blastRequested = QtCore.pyqtSignal(list, dict)   # files, config dict
    stopRequested  = QtCore.pyqtSignal()             # user clicked Stop

    _DATABASES        = ["core_nt", "nt", "refseq_rna", "16S_ribosomal_RNA"]
    _PROGRAMS         = ["blastn&MEGABLAST=on", "blastn", "megablast"]
    _PROGRAM_LABELS   = ["blastn + MEGABLAST (recommended)", "blastn", "megablast"]

    # Live-log slot keys (fixed lines, updated in place)
    _SLOT_KEYS = ("info", "blast", "organism", "taxonomy", "progress", "result")

    def __init__(self, parent=None):
        super().__init__(parent)

        outer_layout = QtWidgets.QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)

        self._inner = QtWidgets.QWidget()
        self._layout = QtWidgets.QVBoxLayout(self._inner)
        self._layout.setContentsMargins(20, 20, 20, 8)
        self._layout.setSpacing(14)
        scroll.setWidget(self._inner)
        outer_layout.addWidget(scroll, 0)   # scroll takes only what content needs

        # ── Title + description ──
        self._lbl_title = make_label("BLAST NCBI Search", size=19, bold=True)
        self._lbl_desc = make_label(
            "BLAST sequences in NCBI. "
            "Drag-and-drop one or more FASTA files (.fa, .fas, .fasta).\n"
            "Optional: results might include organism and taxonomic classification from NCBI Taxonomy.",
            color=TEXT_SEC
        )
        self._lbl_desc.setWordWrap(True)
        self._layout.addWidget(self._lbl_title)
        self._layout.addWidget(self._lbl_desc)

        # ── Settings group ──
        self._settings_box = QtWidgets.QGroupBox("BLAST Settings")
        self._settings_box.setStyleSheet("QGroupBox { font-weight:600; color:#1A1A2E; }")
        sg = QtWidgets.QFormLayout(self._settings_box)
        sg.setLabelAlignment(QtCore.Qt.AlignRight)
        sg.setSpacing(10)
        sg.setContentsMargins(16, 16, 16, 16)

        self._lbl_api = QtWidgets.QLabel("NCBI API Key:")
        api_row = QtWidgets.QWidget()
        al = QtWidgets.QHBoxLayout(api_row)
        al.setContentsMargins(0, 0, 0, 0)
        al.setSpacing(4)
        self._api_key_edit = QtWidgets.QLineEdit()
        self._api_key_edit.setPlaceholderText("Paste your NCBI API key here…")
        self._api_key_edit.setMinimumWidth(500)
        self._api_key_edit.setText(self._load_api_key())
        self._api_key_edit.editingFinished.connect(self._save_api_key)
        self._api_key_edit.textChanged.connect(self._on_api_key_changed)
        al.addWidget(self._api_key_edit)
        self._api_key_clear_btn = QtWidgets.QPushButton("✕")
        self._api_key_clear_btn.setFixedSize(28, 28)
        self._api_key_clear_btn.setToolTip("Clear API key")
        self._api_key_clear_btn.setStyleSheet(
            "QPushButton { background: transparent; border: 1px solid #CCC;"
            " border-radius: 4px; color: #888; font-size:11px; }"
            "QPushButton:hover { background: #FEE2E2; border-color: #EF4444; color: #EF4444; }"
        )
        self._api_key_clear_btn.clicked.connect(self._clear_api_key)
        al.addWidget(self._api_key_clear_btn)
        al.addStretch()
        sg.addRow(self._lbl_api, api_row)

        self._lbl_api_warn = QtWidgets.QLabel(
            "<table cellspacing='0' cellpadding='0'><tr>"
            "<td valign='top'>⚠&nbsp;&nbsp;</td>"
            "<td>A personal NCBI API key is required — key is saved automatically.<br>"
            "Register free at: "
            "<a href='https://www.ncbi.nlm.nih.gov/account' "
            "style='color:#185FA5;'>ncbi.nlm.nih.gov/account</a></td>"
            "</tr></table>"
        )
        self._lbl_api_warn.setWordWrap(True)
        self._lbl_api_warn.setOpenExternalLinks(True)
        self._lbl_api_warn.setStyleSheet("color:#B45309; font-size:16px;")
        sg.addRow("", self._lbl_api_warn)
        self._lbl_api_warn.setVisible(not bool(self._api_key_edit.text().strip()))

        self._db_combo = QtWidgets.QComboBox()
        for db in self._DATABASES:
            self._db_combo.addItem(db)
        self._db_combo.setFixedWidth(500)
        self._lbl_db = QtWidgets.QLabel("Database:")
        sg.addRow(self._lbl_db, self._db_combo)

        self._prog_combo = QtWidgets.QComboBox()
        for lbl in self._PROGRAM_LABELS:
            self._prog_combo.addItem(lbl)
        self._prog_combo.setFixedWidth(500)
        self._lbl_prog = QtWidgets.QLabel("Program:")
        sg.addRow(self._lbl_prog, self._prog_combo)

        hits_row = QtWidgets.QWidget()
        hl = QtWidgets.QHBoxLayout(hits_row)
        hl.setContentsMargins(0, 0, 0, 0)
        self._hits_spin = QtWidgets.QSpinBox()
        self._hits_spin.setRange(1, 100)
        self._hits_spin.setValue(5)
        self._hits_spin.setFixedWidth(80)
        hl.addWidget(self._hits_spin)
        hl.addStretch()
        self._lbl_hits = QtWidgets.QLabel("Hits per sequence (1–100):")
        sg.addRow(self._lbl_hits, hits_row)

        batch_row = QtWidgets.QWidget()
        bl2 = QtWidgets.QHBoxLayout(batch_row)
        bl2.setContentsMargins(0, 0, 0, 0)
        bl2.setSpacing(6)
        self._batch_spin = QtWidgets.QSpinBox()
        self._batch_spin.setRange(1, 100)
        self._batch_spin.setValue(50)
        self._batch_spin.setFixedWidth(80)
        bl2.addWidget(self._batch_spin)
        self._ncbi_warn_icon = QtWidgets.QLabel("⚠")
        self._ncbi_warn_icon.setStyleSheet(
            "color: #B45309; font-size: 17px; padding: 0 2px;"
        )
        self._ncbi_warn_icon.setToolTip(
            "<b>NCBI usage policy warning</b><br><br>"
            "NCBI monitors and penalizes excessive use of its servers.<br>"
            "Submitting too many requests in a short period may result in:<br>"
            "• Temporary or permanent IP blocking<br>"
            "• Suspension of your API key<br><br>"
            "Keep batches ≤ 50 sequences and avoid running multiple<br>"
            "simultaneous BLAST sessions or long sessions in NCBI."
        )
        self._ncbi_warn_icon.setCursor(QtCore.Qt.WhatsThisCursor)
        bl2.addWidget(self._ncbi_warn_icon)
        bl2.addStretch()
        self._lbl_batch = QtWidgets.QLabel("Sequences per batch (max 50 recommended):")
        sg.addRow(self._lbl_batch, batch_row)

        self._tax_check = QtWidgets.QCheckBox("Fetch organism + taxonomic classification")
        self._tax_check.setChecked(True)
        self._lbl_tax = QtWidgets.QLabel("Taxonomy lookup:")
        sg.addRow(self._lbl_tax, self._tax_check)

        self._layout.addWidget(self._settings_box)

        # ── Drop zone ──
        self._drop = MultiDropZone()
        self._drop.filesDropped.connect(self._on_files)
        self._layout.addWidget(self._drop)
        self._layout.addStretch()   # packs content at top; log lives outside the scroll

        # ── Live progress display (outside scroll so it expands to fill space) ──
        self._log_slots = {k: "" for k in self._SLOT_KEYS}
        self._log = QtWidgets.QPlainTextEdit()
        self._log.setReadOnly(True)
        self._log.setFont(QtGui.QFont("Consolas", 9))
        self._log.setMinimumHeight(200)
        self._log.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed
        )
        self._log.setStyleSheet(
            f"QPlainTextEdit {{ background:{GRAY_BG}; border:1px solid {GRAY_LINE}; "
            f"border-radius:6px; padding:6px; color:{TEXT_PRI}; margin:0 20px 8px 20px; "
            f"font-family:'Consolas','Courier New',monospace; }}"
        )
        self._log.hide()
        outer_layout.addWidget(self._log, 0)   # stretch=0: fixed height (~30% of panel, min 200px)

        # ── Elapsed-time timer ──
        self._info_base   = ""    # static part of the info line (without time)
        self._start_time  = 0.0
        self._elapsed_timer = QtCore.QTimer(self)
        self._elapsed_timer.setInterval(1000)
        self._elapsed_timer.timeout.connect(self._tick_elapsed)

        # ── Footer ──
        footer = QtWidgets.QWidget()
        footer.setObjectName("blast_footer")
        footer.setStyleSheet(f"""
            QWidget#blast_footer {{
                background: {GRAY_CARD};
                border-top: 1px solid {GRAY_LINE};
            }}
        """)
        fl = QtWidgets.QHBoxLayout(footer)
        fl.setContentsMargins(20, 10, 20, 10)

        self._clear_btn = QtWidgets.QPushButton("Clear")
        self._clear_btn.setObjectName("danger_btn")
        self._clear_btn.setFixedHeight(44)
        self._clear_btn.setFixedWidth(140)
        self._clear_btn.clicked.connect(self._reset)
        fl.addWidget(self._clear_btn)

        self._open_folder_btn = QtWidgets.QPushButton("Open folder  📂")
        self._open_folder_btn.setObjectName("secondary_btn")
        self._open_folder_btn.setFixedHeight(44)
        self._open_folder_btn.hide()
        self._open_folder_btn.clicked.connect(self._open_output_folder)
        fl.addWidget(self._open_folder_btn)

        self._open_results_btn = QtWidgets.QPushButton("Open results  📄")
        self._open_results_btn.setObjectName("secondary_btn")
        self._open_results_btn.setFixedHeight(44)
        self._open_results_btn.hide()
        self._open_results_btn.clicked.connect(self._open_results_file)
        fl.addWidget(self._open_results_btn)

        self._stop_btn = QtWidgets.QPushButton("Stop")
        self._stop_btn.setObjectName("danger_btn")
        self._stop_btn.setFixedHeight(44)
        self._stop_btn.setFixedWidth(120)
        self._stop_btn.hide()
        self._stop_btn.clicked.connect(self._stop_blast)
        fl.addWidget(self._stop_btn)

        fl.addStretch()

        self._blast_btn = QtWidgets.QPushButton("Run BLAST  →")
        self._blast_btn.setObjectName("primary_btn")
        self._blast_btn.setFixedHeight(44)
        self._blast_btn.setFixedWidth(300)
        self._blast_btn.setEnabled(False)
        self._blast_btn.clicked.connect(self._emit_blast)
        self._blast_btn.setStyleSheet(
            f"QPushButton {{ background-color: {GRAY_LINE}; color: {TEXT_HINT}; border:none; "
            f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
        )
        fl.addWidget(self._blast_btn)
        outer_layout.addWidget(footer)

        self._last_outdir = ""
        self._last_tsv    = ""
        self._worker: Optional[_BlastWorker] = None
        # Workers detached but possibly still running — kept referenced so they are
        # never garbage-collected mid-run (which crashes Qt). Purged in _retire_worker.
        self._retired_workers: set = set()

        self.installEventFilter(self)

    def eventFilter(self, obj, event):
        """Detectar cuando el panel cambia de tamaño para ajustar el log."""
        if obj == self and event.type() == QtCore.QEvent.Resize:
            self._adjust_log_height()
        return super().eventFilter(obj, event)

    def _adjust_log_height(self):
        """Ajustar la altura del log al 30% del panel (mínimo 200px)."""
        if self._log.isVisible():
            target_height = max(int(self.height() * 0.3), 200)
            self._log.setFixedHeight(target_height)

    def showEvent(self, event):
        """Cuando el panel se muestra por primera vez."""
        super().showEvent(event)
        self._adjust_log_height()

    # ── API key persistence ───────────────────────────────────────────────

    @staticmethod
    def _config_path():
        return os.path.join(_profiles_dir(), "blast_config.json")

    def _load_api_key(self) -> str:
        try:
            with open(self._config_path(), "r", encoding="utf-8") as f:
                return _json_mod.load(f).get("api_key", "")
        except Exception:
            return ""

    def _save_api_key(self):
        path = self._config_path()
        try:
            data: dict = {}
            if os.path.isfile(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        data = _json_mod.load(f)
                except Exception:
                    pass
            data["api_key"] = self._api_key_edit.text().strip()
            with open(path, "w", encoding="utf-8") as f:
                _json_mod.dump(data, f, indent=2)
        except Exception:
            pass

    def _on_api_key_changed(self, text: str):
        self._lbl_api_warn.setVisible(not bool(text.strip()))

    def _clear_api_key(self):
        self._api_key_edit.clear()
        self._save_api_key()

    # ── UI helpers ────────────────────────────────────────────────────────

    def retranslateUi(self):
        ctx = "BlastPanel"
        self._lbl_title.setText(_tr(ctx, "BLAST NCBI Search"))
        self._lbl_desc.setText(_tr(ctx,
            "BLAST multiFASTA sequences against NCBI. "
            "Drag-and-drop one or more FASTA files (.fa, .fas, .fasta). "
            "Results include top hits with organism and taxonomic classification."))
        self._settings_box.setTitle(_tr(ctx, "BLAST Settings"))
        self._lbl_api.setText(_tr(ctx, "NCBI API Key:"))
        self._lbl_api_warn.setText(_tr(ctx,
            "⚠  A personal NCBI API key is required. Register free at: "
            "ncbi.nlm.nih.gov/account — key is saved automatically."))
        self._lbl_db.setText(_tr(ctx, "Database:"))
        self._lbl_prog.setText(_tr(ctx, "Program:"))
        self._lbl_hits.setText(_tr(ctx, "Hits per sequence (1–100):"))
        self._lbl_batch.setText(_tr(ctx, "Sequences per batch (max 50 recommended):"))
        self._lbl_tax.setText(_tr(ctx, "Taxonomy lookup:"))
        self._tax_check.setText(_tr(ctx, "Fetch organism + taxonomy"))
        self._clear_btn.setText(_tr(ctx, "Clear"))
        self._open_folder_btn.setText(_tr(ctx, "Open folder  📂"))
        self._open_results_btn.setText(_tr(ctx, "Open results  📄"))
        self._blast_btn.setText(_tr(ctx, "Run BLAST  →"))
        self._drop.retranslateUi()

    def changeEvent(self, event):
        if event.type() == QtCore.QEvent.LanguageChange:
            self.retranslateUi()
        super().changeEvent(event)

    # ── Slots ─────────────────────────────────────────────────────────────

    def _on_files(self, paths):
        enabled = len(paths) >= 1
        self._blast_btn.setEnabled(enabled)
        if enabled:
            self._blast_btn.setStyleSheet(
                f"QPushButton {{ background-color: {BLUE}; color: white; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
                f"QPushButton:hover {{ background-color: #0C4A82; }}"
            )
        else:
            self._blast_btn.setStyleSheet(
                f"QPushButton {{ background-color: {GRAY_LINE}; color: {TEXT_HINT}; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
            )

    def _emit_blast(self):
        api_key = self._api_key_edit.text().strip()
        if not api_key:
            QtWidgets.QMessageBox.warning(
                self, "NCBI API Key required",
                "Please enter a valid NCBI API key before running BLAST."
            )
            return

        ts          = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"ont-barcoder_{ts}_blast"
        auto_dir    = os.path.join(_get_base_dir(), "output", folder_name)

        out_dir = self._ask_output_dir(auto_dir, folder_name)
        if out_dir is None:
            return

        self._save_api_key()

        cfg = {
            "api_key":        api_key,
            "database":       self._DATABASES[self._db_combo.currentIndex()],
            "program":        self._PROGRAMS[self._prog_combo.currentIndex()],
            "nhits":          self._hits_spin.value(),
            "nseq":           self._batch_spin.value(),
            "fetch_taxonomy": self._tax_check.isChecked(),
            "outdir":         out_dir,
        }

        self._retire_worker()

        self._worker = _BlastWorker(list(self._drop.files), cfg)
        self._worker.statusUpdated.connect(self.update_status)
        self._worker.progressUpdated.connect(self.set_progress)
        self._worker.taskFinished.connect(self.on_finished)
        self._worker.taskError.connect(self.on_error)
        self.set_running(True)
        self._worker.start()

    def _stop_blast(self):
        if self._worker and self._worker.isRunning():
            self._worker.stop()

    def _retire_worker(self):
        """Detach the current worker so a late signal can't touch the UI, and keep a
        reference until its thread actually exits (a QThread garbage-collected while
        still running crashes Qt). Stopping does not interrupt a blocked urlopen, so
        we never wait() here — the thread is retained and released on the next call."""
        self._retired_workers = {w for w in self._retired_workers if w.isRunning()}
        w = self._worker
        self._worker = None
        if w is None:
            return
        for sig in (w.statusUpdated, w.progressUpdated, w.taskFinished, w.taskError):
            try:
                sig.disconnect()
            except (TypeError, RuntimeError):
                pass
        if w.isRunning():
            w.stop()
            self._retired_workers.add(w)

    def _ask_output_dir(self, auto_dir: str, folder_name: str) -> Optional[str]:
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Output folder")
        dlg.setMinimumWidth(480)
        dlg.setStyleSheet(f"""
            QDialog {{ background-color: {GRAY_CARD}; }}
            QLabel {{ color: {TEXT_PRI}; background-color: transparent; }}
            QRadioButton {{
                color: {TEXT_PRI}; background-color: transparent;
                font-size: 15px; padding: 6px 0;
            }}
            QRadioButton::indicator {{ width: 16px; height: 16px; }}
            QPushButton {{
                border-radius: 8px; padding: 8px 20px;
                font-size: 15px; font-weight: 500;
            }}
            #dlg_ok_btn {{ background-color: {BLUE}; color: white; border: none; }}
            #dlg_ok_btn:hover {{ background-color: #0C4A82; }}
            #dlg_cancel_btn {{
                background-color: transparent; color: {BLUE};
                border: 1px solid {BLUE};
            }}
            #dlg_cancel_btn:hover {{ background-color: {BLUE_LIGHT}; }}
        """)

        vlay = QtWidgets.QVBoxLayout(dlg)
        vlay.setSpacing(16)
        vlay.setContentsMargins(24, 24, 24, 20)

        title_lbl = QtWidgets.QLabel("Where to save the results?")
        title_lbl.setStyleSheet(f"font-size:17px; font-weight:700; color:{TEXT_PRI};")
        vlay.addWidget(title_lbl)

        radio_auto = QtWidgets.QRadioButton(
            f"Automatic folder (recommended)\n  …/output/{folder_name}/"
        )
        radio_auto.setChecked(True)
        radio_custom = QtWidgets.QRadioButton("Select folder manually")
        vlay.addWidget(radio_auto)
        vlay.addWidget(radio_custom)
        vlay.addSpacing(8)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_cancel.setObjectName("dlg_cancel_btn")
        btn_cancel.setFixedHeight(38)
        btn_ok = QtWidgets.QPushButton("Run")
        btn_ok.setObjectName("dlg_ok_btn")
        btn_ok.setFixedHeight(38)
        btn_ok.setDefault(True)
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addSpacing(8)
        btn_row.addWidget(btn_ok)
        vlay.addLayout(btn_row)

        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return None

        if radio_auto.isChecked():
            return auto_dir
        parent_dir = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select output folder"
        )
        if not parent_dir:
            return None
        return os.path.join(parent_dir, folder_name)

    def _open_output_folder(self):
        if self._last_outdir and os.path.isdir(self._last_outdir):
            os.startfile(self._last_outdir)

    def _open_results_file(self):
        if self._last_tsv and os.path.isfile(self._last_tsv):
            os.startfile(self._last_tsv)

    def _reset(self):
        self._drop.clear()
        for k in self._SLOT_KEYS:
            self._log_slots[k] = ""
        self._info_base = ""
        self._elapsed_timer.stop()
        self._log.clear()
        self._log.hide()
        self._open_folder_btn.hide()
        self._open_results_btn.hide()
        self._stop_btn.hide()
        self._blast_btn.show()
        self._blast_btn.setEnabled(False)
        self._blast_btn.setStyleSheet(
            f"QPushButton {{ background-color: {GRAY_LINE}; color: {TEXT_HINT}; border:none; "
            f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
        )
        self._last_outdir = ""
        self._last_tsv    = ""
        self._clear_btn.setEnabled(True)

    # ── Public API (called by MainWindow) ──────────────────────────────────

    def _rebuild_log(self):
        """Rewrite the entire log widget from the fixed slots dict."""
        sep = "─" * 56
        lines = [
            self._log_slots.get("info",     ""),
            sep,
            self._log_slots.get("blast",    ""),
            self._log_slots.get("organism", ""),
            self._log_slots.get("taxonomy", ""),
            self._log_slots.get("progress", ""),
            sep,
            self._log_slots.get("result",   ""),
        ]
        self._log.setPlainText("\n".join(lines))

    def update_status(self, key: str, text: str):
        """Update a named live slot and refresh the display."""
        self._log.show()
        self._adjust_log_height()
        if key == "info":
            self._info_base = text      # save static part; timer appends elapsed
        self._log_slots[key] = text
        self._rebuild_log()

    def _tick_elapsed(self):
        """Called every second by the timer to update elapsed time in the info line."""
        elapsed = int(time.monotonic() - self._start_time)
        h, rem  = divmod(elapsed, 3600)
        m, s    = divmod(rem, 60)
        t_str   = f"{h}h {m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"
        self._log_slots["info"] = f"{self._info_base}  │  Time: {t_str}"
        self._rebuild_log()

    def set_running(self, running: bool):
        self._blast_btn.setVisible(not running)
        self._stop_btn.setVisible(running)
        self._clear_btn.setEnabled(not running)
        if running:
            self._start_time = time.monotonic()
            self._elapsed_timer.start()
            self._log.show()
            self._adjust_log_height()
        else:
            self._elapsed_timer.stop()

    def set_progress(self, current: int, total: int):
        if total > 0:
            pct = int(current * 100 / total)
            bar_len = 28
            filled = int(bar_len * current / total)
            bar = "█" * filled + " " * (bar_len - filled)
            self.update_status(
                "progress",
                f"Progress    │ [{bar}] {pct}%"
            )

    def on_finished(self, outdir: str):
        self.set_running(False)
        self._last_outdir = outdir
        if outdir and os.path.isdir(outdir):
            self._open_folder_btn.show()
            # Prefer xlsx; fall back to tsv if xlsx conversion failed
            for ext in (".xlsx", ".tsv"):
                matches = sorted(
                    (os.path.join(outdir, f) for f in os.listdir(outdir)
                     if f.endswith(ext) and f.startswith("blast-")),
                    key=os.path.getmtime, reverse=True
                )
                if matches:
                    self._last_tsv = matches[0]
                    self._open_results_btn.show()
                    break
        # Do NOT overwrite the result slot — the worker already set the final message
        self._blast_btn.setEnabled(bool(self._drop.files))
        if self._drop.files:
            self._blast_btn.setStyleSheet(
                f"QPushButton {{ background-color: {BLUE}; color: white; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
                f"QPushButton:hover {{ background-color: #0C4A82; }}"
            )

    def on_error(self, msg: str):
        self.set_running(False)
        self.update_status("result", f"ERROR       │ {msg[:80]}")
        self._blast_btn.setEnabled(bool(self._drop.files))
        if self._drop.files:
            self._blast_btn.setStyleSheet(
                f"QPushButton {{ background-color: {BLUE}; color: white; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
                f"QPushButton:hover {{ background-color: #0C4A82; }}"
            )


# ═══════════════════════════════════════════════════════════════════════════
# LIVE CHART PANEL
# ═══════════════════════════════════════════════════════════════════════════


class _FastqInspectorWorker(QtCore.QThread):
    progress = QtCore.pyqtSignal(int)   # reads processed so far
    finished = QtCore.pyqtSignal(dict)
    error    = QtCore.pyqtSignal(str)

    _SCATTER_MAX = 10_000

    def __init__(self, path: str, parent=None):
        super().__init__(parent)
        self._path = path
        self._stop = False

    def stop(self):
        self._stop = True

    # ── minimum file size to justify spawning worker processes ──────────────
    _PARALLEL_THRESHOLD = 50 * 1024 * 1024   # 50 MB uncompressed

    def run(self):
        try:
            path  = self._path
            is_gz = path.lower().endswith(".gz")
            size  = os.path.getsize(path)
            # Gzip streams cannot be split; only parallelize plain FASTQ files
            # large enough to amortize process-spawn overhead (~200 ms on Windows).
            if not is_gz and size >= self._PARALLEL_THRESHOLD:
                self._run_parallel(path, size)
            else:
                self._run_sequential(path, is_gz)
        except Exception as exc:
            self.error.emit(str(exc))

    def _run_sequential(self, path: str, is_gz: bool):
        """Read all records in one thread.  Used for gzip files and small plain FASTQ."""
        import gzip, io

        lengths:  List[int]   = []
        q_scores: List[float] = []
        gc_pcts:  List[float] = []
        total_bases = 0
        q_sum = gc_sum = 0.0
        len_lt500 = len_500_2k = len_gt2k = 0
        q_cnt_10 = q_cnt_15 = q_cnt_20 = 0

        # Open in binary mode so quality bytes can be summed directly (no ord() call).
        # For gzip, wrap the raw file in a 4 MB BufferedReader so compressed data
        # is read in large chunks before decompression — reduces I/O syscall overhead.
        if is_gz:
            raw = open(path, "rb", buffering=4 * 1024 * 1024)
            fh  = gzip.GzipFile(fileobj=raw)
        else:
            fh  = open(path, "rb", buffering=4 * 1024 * 1024)
            raw = fh

        try:
            it = iter(fh)
            for header in it:
                if self._stop:
                    return
                # Skip malformed / mid-record lines so this path matches the parallel
                # one (_fq_chunk_task), which also resyncs on non-'@' headers.
                if not header.startswith(b"@"):
                    continue
                try:
                    seq  = next(it).rstrip(b"\r\n")
                    next(it)                           # '+' line
                    qual = next(it).rstrip(b"\r\n")
                except StopIteration:
                    break

                L  = len(seq)
                ql = len(qual) or 1
                if L == 0:
                    continue

                # C-speed GC count — str.count() / bytes.count() implemented in C
                gc_cnt = (seq.count(b"G") + seq.count(b"C") +
                          seq.count(b"g") + seq.count(b"c"))
                gc     = gc_cnt / L * 100.0

                # Fast Q-mean — bytes iterate as integers; no ord() needed
                q_mean = (sum(qual) - 33 * ql) / ql

                lengths.append(L)
                q_scores.append(q_mean)
                gc_pcts.append(gc)

                total_bases += L
                q_sum       += q_mean
                gc_sum      += gc

                if L < 500:      len_lt500  += 1
                elif L <= 2000:  len_500_2k += 1
                else:            len_gt2k   += 1

                if q_mean >= 10: q_cnt_10 += 1
                if q_mean >= 15: q_cnt_15 += 1
                if q_mean >= 20: q_cnt_20 += 1

                if len(lengths) % 5000 == 0:
                    self.progress.emit(len(lengths))
        finally:
            fh.close()
            if is_gz:
                raw.close()

        if self._stop:
            return

        self._compute_and_emit(
            lengths, q_scores, gc_pcts,
            total_bases, q_sum, gc_sum,
            len_lt500, len_500_2k, len_gt2k,
            q_cnt_10, q_cnt_15, q_cnt_20,
        )

    def _run_parallel(self, path: str, file_size: int):
        """Split an uncompressed FASTQ file into chunks and process them in parallel
        using ProcessPoolExecutor.  Each worker runs _fq_chunk_task() in its own
        Python process, bypassing the GIL for true CPU-level parallelism."""
        import random
        from concurrent.futures import ProcessPoolExecutor, as_completed

        n_workers = min(os.cpu_count() or 1, 8)
        chunk     = file_size // n_workers

        # Build (start, end) byte ranges — end=0 means "read until EOF"
        boundaries = [(i * chunk, (i + 1) * chunk if i < n_workers - 1 else 0)
                      for i in range(n_workers)]

        all_lengths:  List[int]   = []
        all_q_scores: List[float] = []
        all_gc_pcts:  List[float] = []
        total_bases = 0
        q_sum = gc_sum = 0.0
        len_lt500 = len_500_2k = len_gt2k = 0
        q_cnt_10 = q_cnt_15 = q_cnt_20 = 0

        with ProcessPoolExecutor(max_workers=n_workers) as pool:
            futures = {pool.submit(_fq_chunk_task, path, s, e): i
                       for i, (s, e) in enumerate(boundaries)}
            for future in as_completed(futures):
                if self._stop:
                    pool.shutdown(wait=False, cancel_futures=True)
                    return
                p = future.result()
                all_lengths.extend(p["lengths"])
                all_q_scores.extend(p["q_scores"])
                all_gc_pcts.extend(p["gc_pcts"])
                total_bases  += p["total_bases"]
                q_sum        += p["q_sum"]
                gc_sum       += p["gc_sum"]
                len_lt500    += p["len_lt500"]
                len_500_2k   += p["len_500_2k"]
                len_gt2k     += p["len_gt2k"]
                q_cnt_10     += p["q_cnt_10"]
                q_cnt_15     += p["q_cnt_15"]
                q_cnt_20     += p["q_cnt_20"]
                self.progress.emit(len(all_lengths))

        if self._stop:
            return

        self._compute_and_emit(
            all_lengths, all_q_scores, all_gc_pcts,
            total_bases, q_sum, gc_sum,
            len_lt500, len_500_2k, len_gt2k,
            q_cnt_10, q_cnt_15, q_cnt_20,
        )

    def _compute_and_emit(
        self,
        lengths, q_scores, gc_pcts,
        total_bases, q_sum, gc_sum,
        len_lt500, len_500_2k, len_gt2k,
        q_cnt_10, q_cnt_15, q_cnt_20,
    ):
        """Compute final statistics and histogram bins, then emit finished signal.
        Shared by both sequential and parallel paths."""
        import math, random

        N_BINS = 60
        n = len(lengths)
        if n == 0:
            self.error.emit("No reads found in file.")
            return

        mean_len    = total_bases / n
        sorted_lens = sorted(lengths)
        median_len  = (sorted_lens[n // 2 - 1] + sorted_lens[n // 2]) / 2 \
                      if n % 2 == 0 else sorted_lens[n // 2]
        min_len = sorted_lens[0]
        max_len = sorted_lens[-1]

        # N50
        half, cumsum, n50 = total_bases / 2, 0, sorted_lens[-1]
        for l in reversed(sorted_lens):
            cumsum += l
            if cumsum >= half:
                n50 = l
                break

        mean_q  = q_sum  / n
        mean_gc = gc_sum / n
        std_gc  = math.sqrt(sum((x - mean_gc) ** 2 for x in gc_pcts) / max(n - 1, 1))

        len_p01 = sorted_lens[max(0,     int(n * 0.01))]
        len_p99 = sorted_lens[min(n - 1, int(n * 0.99))]

        def _make_bins(values, v_min, v_max):
            span = v_max - v_min if v_max > v_min else 1.0
            bins = [0] * N_BINS
            for v in values:
                if v_min <= v <= v_max:
                    bins[min(int((v - v_min) / span * N_BINS), N_BINS - 1)] += 1
            return bins

        # Full-data Q range — the histogram bins are built over this range, so the
        # chart axis must use the SAME bounds (not the sampled scatter subset below,
        # whose min/max can be narrower and would misalign the bars).
        q_min_full = min(q_scores)
        q_max_full = max(q_scores)

        len_bins = _make_bins(lengths,  len_p01, len_p99)
        q_bins   = _make_bins(q_scores, q_min_full, q_max_full)
        gc_bins  = _make_bins(gc_pcts,  0.0, 100.0)

        if n > self._SCATTER_MAX:
            idx    = random.sample(range(n), self._SCATTER_MAX)
            sc_len = [lengths[i]  for i in idx]
            sc_q   = [q_scores[i] for i in idx]
        else:
            sc_len = lengths[:]
            sc_q   = q_scores[:]

        del lengths, q_scores, gc_pcts, sorted_lens

        self.finished.emit({
            "n_reads":    n,
            "total_bases": total_bases,
            "mean_len":   mean_len,
            "median_len": median_len,
            "min_len":    min_len,
            "max_len":    max_len,
            "n50":        n50,
            "len_lt500":  len_lt500,
            "len_500_2k": len_500_2k,
            "len_gt2k":   len_gt2k,
            "mean_q":     mean_q,
            "q10_pct":    q_cnt_10 / n * 100,
            "q15_pct":    q_cnt_15 / n * 100,
            "q20_pct":    q_cnt_20 / n * 100,
            "mean_gc":    mean_gc,
            "std_gc":     std_gc,
            "len_bins":   len_bins,
            "q_bins":     q_bins,
            "gc_bins":    gc_bins,
            "q_min":      q_min_full,
            "q_max":      q_max_full,
            "sc_len":     sc_len,
            "sc_q":       sc_q,
            "len_p01":    len_p01,
            "len_p99":    len_p99,
        })


# ═══════════════════════════════════════════════════════════════════════════
# FASTQ INSPECTOR – panel
# ═══════════════════════════════════════════════════════════════════════════

class FastqInspectorPanel(QtWidgets.QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)

        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── Scroll area ──────────────────────────────────────────────────
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        inner = QtWidgets.QWidget()
        self._layout = QtWidgets.QVBoxLayout(inner)
        self._layout.setContentsMargins(20, 20, 20, 20)
        self._layout.setSpacing(16)
        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        # ── Title ──
        self._layout.addWidget(make_label("FASTQ Inspector", size=19, bold=True))
        desc = make_label(
            "Drag a FASTQ file (.fastq or .fastq.gz) to get a full descriptive report. "
            "All reads are analyzed; only the scatter plot uses a random sample of 10 000 points.",
            color=TEXT_SEC,
        )
        desc.setWordWrap(True)
        self._layout.addWidget(desc)

        # ── Drop zone ──
        self._drop = DropZone(
            "Drag FASTQ file here",
            "Accepted: .fastq  |  .fastq.gz",
            extensions=[".fastq", ".gz"],
        )
        self._drop.fileDropped.connect(self._on_file)
        self._layout.addWidget(self._drop)

        # ── Status + progress bar ──
        self._status_lbl = make_label("", color=TEXT_SEC)
        self._layout.addWidget(self._status_lbl)

        self._progress_bar = QtWidgets.QProgressBar()
        self._progress_bar.setRange(0, 0)
        self._progress_bar.setFixedHeight(6)
        self._progress_bar.hide()
        self._layout.addWidget(self._progress_bar)

        # ── Stats section ──
        # Local helper: section label in BLUE for this panel
        def _sec(text):
            lbl = make_section_label(text)
            lbl.setStyleSheet(
                f"font-size:18px; font-weight:600; color:{BLUE}; letter-spacing:0.5px;"
            )
            return lbl

        self._stats_w = QtWidgets.QWidget()
        self._stats_w.hide()
        sl = QtWidgets.QVBoxLayout(self._stats_w)
        sl.setContentsMargins(0, 0, 0, 0)
        sl.setSpacing(10)

        sl.addWidget(_sec("General"))
        r = QtWidgets.QHBoxLayout()
        r.setSpacing(8)
        self._sc_reads    = self._card("Total reads")
        self._sc_bases    = self._card("Total bases")
        self._sc_filesize = self._card("File size")
        for c in (self._sc_reads, self._sc_bases, self._sc_filesize):
            r.addWidget(c)
        r.addStretch()
        sl.addLayout(r)

        sl.addWidget(_sec("Read length"))
        r2 = QtWidgets.QHBoxLayout()
        r2.setSpacing(8)
        self._sc_min    = self._card("Min")
        self._sc_max    = self._card("Max")
        self._sc_mean   = self._card("Mean")
        self._sc_median = self._card("Median")
        self._sc_n50    = self._card("N50", color=BLUE)
        for c in (self._sc_min, self._sc_max, self._sc_mean, self._sc_median, self._sc_n50):
            r2.addWidget(c)
        sl.addLayout(r2)

        r3 = QtWidgets.QHBoxLayout()
        r3.setSpacing(8)
        self._sc_lt500  = self._card("< 500 bp")
        self._sc_500_2k = self._card("500–2000 bp")
        self._sc_gt2k   = self._card("> 2000 bp")
        for c in (self._sc_lt500, self._sc_500_2k, self._sc_gt2k):
            r3.addWidget(c)
        r3.addStretch(2)
        sl.addLayout(r3)

        sl.addWidget(_sec("Base quality (Phred)"))
        r4 = QtWidgets.QHBoxLayout()
        r4.setSpacing(8)
        self._sc_meanq = self._card("Mean Q-score", color=GREEN)
        self._sc_q10   = self._card("Reads Q≥10")
        self._sc_q15   = self._card("Reads Q≥15")
        self._sc_q20   = self._card("Reads Q≥20")
        for c in (self._sc_meanq, self._sc_q10, self._sc_q15, self._sc_q20):
            r4.addWidget(c)
        r4.addStretch()
        sl.addLayout(r4)

        sl.addWidget(_sec("GC content"))
        r5 = QtWidgets.QHBoxLayout()
        r5.setSpacing(8)
        self._sc_meangc = self._card("Mean %GC")
        self._sc_stdgc  = self._card("Std dev %GC")
        for c in (self._sc_meangc, self._sc_stdgc):
            r5.addWidget(c)
        r5.addStretch(3)
        sl.addLayout(r5)

        self._layout.addWidget(self._stats_w)

        # ── Charts section ──
        self._charts_w = QtWidgets.QWidget()
        self._charts_w.hide()
        cl = QtWidgets.QVBoxLayout(self._charts_w)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(12)

        cl.addWidget(_sec("Distributions"))

        row_c1 = QtWidgets.QHBoxLayout()
        row_c1.setSpacing(12)
        self._hist_len = _HistWidget("Read length distribution", "Length (bp)",  color=BLUE_MID)
        self._hist_q   = _HistWidget("Mean Q-score per read",    "Mean Q-score", color=GREEN_MID)
        row_c1.addWidget(self._hist_len)
        row_c1.addWidget(self._hist_q)
        cl.addLayout(row_c1)

        row_c2 = QtWidgets.QHBoxLayout()
        row_c2.setSpacing(12)
        self._hist_gc = _HistWidget("GC content per read",              "% GC",        color=AMBER)
        self._scatter  = _ScatterWidget("Read length vs. mean Q-score", "Length (bp)", "Mean Q-score", color=BLUE_MID)
        row_c2.addWidget(self._hist_gc)
        row_c2.addWidget(self._scatter)
        cl.addLayout(row_c2)

        self._layout.addWidget(self._charts_w)
        self._layout.addStretch()

        # ── Footer ──────────────────────────────────────────────────────
        footer = QtWidgets.QWidget()
        footer.setObjectName("fqins_footer")
        footer.setStyleSheet(f"""
            QWidget#fqins_footer {{
                background: {GRAY_CARD};
                border-top: 1px solid {GRAY_LINE};
            }}
        """)
        fl = QtWidgets.QHBoxLayout(footer)
        fl.setContentsMargins(20, 10, 20, 10)
        fl.setSpacing(8)

        self._clear_btn = QtWidgets.QPushButton("Clear")
        self._clear_btn.setObjectName("danger_btn")
        self._clear_btn.setFixedHeight(44)
        self._clear_btn.setFixedWidth(140)
        self._clear_btn.clicked.connect(self._clear)
        fl.addWidget(self._clear_btn)

        self._pdf_btn = QtWidgets.QPushButton("Export PDF  ↓")
        self._pdf_btn.setObjectName("secondary_btn")
        self._pdf_btn.setFixedHeight(44)
        self._pdf_btn.setFixedWidth(190)
        self._pdf_btn.hide()
        self._pdf_btn.clicked.connect(self._on_pdf_clicked)
        fl.addWidget(self._pdf_btn)

        fl.addStretch()

        self._analyze_btn = QtWidgets.QPushButton("Analyze  →")
        self._analyze_btn.setObjectName("primary_btn")
        self._analyze_btn.setFixedHeight(44)
        self._analyze_btn.setFixedWidth(300)
        self._analyze_btn.setEnabled(False)
        self._analyze_btn.clicked.connect(self._start)
        self._analyze_btn.setStyleSheet(
            f"QPushButton {{ background-color: {GRAY_LINE}; color: {TEXT_HINT}; border:none; "
            f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
        )
        fl.addWidget(self._analyze_btn)
        outer.addWidget(footer)

        self._worker: Optional[_FastqInspectorWorker] = None
        # Detached-but-maybe-running workers, kept referenced so they aren't GC'd
        # mid-run (which crashes Qt). Purged in _retire_worker.
        self._retired_workers: set = set()
        self._file_path   = ""
        self._file_size   = 0
        self._last_result: Optional[dict] = None
        self._last_pdf    = ""
        self._pdf_ready   = False   # True after first successful export

    # ── helpers ──────────────────────────────────────────────────────────

    def _card(self, label: str, color: str = TEXT_PRI) -> QtWidgets.QFrame:
        frame = QtWidgets.QFrame()
        frame.setObjectName("stat_card")
        vl = QtWidgets.QVBoxLayout(frame)
        vl.setContentsMargins(10, 8, 10, 8)
        vl.setSpacing(2)
        vl.addWidget(make_label(label, size=15, color=TEXT_SEC))
        val_lbl = make_label("—", size=18, bold=True, color=color)
        vl.addWidget(val_lbl)
        frame._val = val_lbl   # type: ignore[attr-defined]
        return frame

    def _set_analyze_enabled(self, enabled: bool):
        self._analyze_btn.setEnabled(enabled)
        if enabled:
            self._analyze_btn.setStyleSheet(
                f"QPushButton {{ background-color: {BLUE}; color: white; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
                f"QPushButton:hover {{ background-color: #0C4A82; }}"
            )
        else:
            self._analyze_btn.setStyleSheet(
                f"QPushButton {{ background-color: {GRAY_LINE}; color: {TEXT_HINT}; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
            )

    # PDF display size per chart (matches HTML img width/height)
    _PDF_CHART_W = 520
    _PDF_CHART_H = 280
    _PDF_SCALE   = 3   # oversample factor → renders at 1110×690, shown at 370×230

    def _chart_to_b64(self, widget) -> str:
        """Render at 3× for crisp fonts, then downscale to the exact display size.
        The embedded PNG is already at _PDF_CHART_W × _PDF_CHART_H so Qt needs
        no scaling and the image occupies exactly the expected space in the PDF."""
        w   = self._PDF_CHART_W * self._PDF_SCALE
        h   = self._PDF_CHART_H * self._PDF_SCALE
        pix = widget.render_to_pixmap(w, h, float(self._PDF_SCALE))
        pix = pix.scaled(self._PDF_CHART_W, self._PDF_CHART_H,
                         QtCore.Qt.IgnoreAspectRatio,
                         QtCore.Qt.SmoothTransformation)
        buf = QtCore.QBuffer()
        buf.open(QtCore.QIODevice.WriteOnly)
        pix.save(buf, "PNG")
        return bytes(buf.data().toBase64()).decode()

    def _retire_worker(self):
        """Detach the current worker: disconnect its slots so stale results never
        reach the UI, and keep a reference until its thread exits so it is never
        garbage-collected while still running (which crashes Qt)."""
        self._retired_workers = {w for w in self._retired_workers if w.isRunning()}
        w = self._worker
        self._worker = None
        if w is None:
            return
        for sig in (w.progress, w.finished, w.error):
            try:
                sig.disconnect()
            except (TypeError, RuntimeError):
                pass
        if w.isRunning():
            w.stop()
            self._retired_workers.add(w)

    # ── slots ─────────────────────────────────────────────────────────────

    def _on_file(self, path: str):
        # Cancel any running analysis — new file supersedes the old one
        self._retire_worker()
        self._file_path = path
        self._set_analyze_enabled(bool(path))
        self._status_lbl.setStyleSheet("")
        self._status_lbl.setText("")
        self._stats_w.hide()
        self._charts_w.hide()
        self._pdf_btn.hide()
        self._pdf_btn.setText("Export PDF  ↓")
        self._pdf_ready = False

    def _clear(self):
        self._retire_worker()
        self._drop.clear()
        self._file_path   = ""
        self._file_size   = 0
        self._last_result = None
        self._last_pdf    = ""
        self._pdf_ready   = False
        self._status_lbl.setStyleSheet("")
        self._status_lbl.setText("")
        self._progress_bar.hide()
        self._stats_w.hide()
        self._charts_w.hide()
        self._pdf_btn.setText("Export PDF  ↓")
        self._pdf_btn.hide()
        self._set_analyze_enabled(False)

    def _start(self):
        if not self._file_path or not os.path.isfile(self._file_path):
            return
        # Stop any previous worker before starting a new one
        self._retire_worker()
        self._file_size = os.path.getsize(self._file_path)
        self._set_analyze_enabled(False)
        self._status_lbl.setStyleSheet("")
        self._status_lbl.setText("Analyzing…  reading all reads")
        self._progress_bar.show()
        self._stats_w.hide()
        self._charts_w.hide()
        self._pdf_btn.hide()

        self._worker = _FastqInspectorWorker(self._file_path)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._worker.start()

    def _on_progress(self, n: int):
        self._status_lbl.setText(f"Analyzing…  {n:,} reads processed")

    def _on_finished(self, r: dict):
        self._progress_bar.hide()
        self._last_result = r
        n = r["n_reads"]
        self._status_lbl.setText(f"Done — {n:,} reads analyzed")

        def _fmt_bp(v):
            if v >= 1_000_000_000: return f"{v/1_000_000_000:.2f} Gbp"
            if v >= 1_000_000:     return f"{v/1_000_000:.1f} Mbp"
            if v >= 1_000:         return f"{v/1_000:.1f} Kbp"
            return f"{v} bp"

        def _fmt_size(b):
            if b >= 1_073_741_824: return f"{b/1_073_741_824:.2f} GB"
            if b >= 1_048_576:     return f"{b/1_048_576:.1f} MB"
            return f"{b/1_024:.1f} KB"

        self._sc_reads._val.setText(f"{n:,}")
        self._sc_bases._val.setText(_fmt_bp(r["total_bases"]))
        self._sc_filesize._val.setText(_fmt_size(self._file_size))

        self._sc_min._val.setText(f"{r['min_len']:,} bp")
        self._sc_max._val.setText(f"{r['max_len']:,} bp")
        self._sc_mean._val.setText(f"{r['mean_len']:,.0f} bp")
        self._sc_median._val.setText(f"{r['median_len']:,.0f} bp")
        self._sc_n50._val.setText(f"{r['n50']:,} bp")

        self._sc_lt500._val.setText(f"{r['len_lt500']:,}  ({r['len_lt500']/n*100:.1f}%)")
        self._sc_500_2k._val.setText(f"{r['len_500_2k']:,}  ({r['len_500_2k']/n*100:.1f}%)")
        self._sc_gt2k._val.setText(f"{r['len_gt2k']:,}  ({r['len_gt2k']/n*100:.1f}%)")

        self._sc_meanq._val.setText(f"Q{r['mean_q']:.1f}")
        self._sc_q10._val.setText(f"{r['q10_pct']:.1f}%")
        self._sc_q15._val.setText(f"{r['q15_pct']:.1f}%")
        self._sc_q20._val.setText(f"{r['q20_pct']:.1f}%")

        self._sc_meangc._val.setText(f"{r['mean_gc']:.1f}%")
        self._sc_stdgc._val.setText(f"± {r['std_gc']:.1f}%")

        # Bins were pre-computed in the worker thread — O(1) on main thread
        self._hist_len.set_precomputed(r["len_bins"], r["len_p01"], r["len_p99"])
        self._hist_q.set_precomputed(r["q_bins"],   r["q_min"],   r["q_max"])
        self._hist_gc.set_precomputed(r["gc_bins"],  0.0,          100.0)
        self._scatter.set_data(r["sc_len"], r["sc_q"],
                               x_min=r["len_p01"], x_max=r["len_p99"],
                               y_min=r["q_min"], y_max=r["q_max"])

        self._stats_w.show()
        self._charts_w.show()
        self._set_analyze_enabled(False)
        self._pdf_btn.setText("Export PDF  ↓")
        self._pdf_ready = False
        self._pdf_btn.show()

    def _on_error(self, msg: str):
        self._progress_bar.hide()
        self._set_analyze_enabled(True)
        self._status_lbl.setStyleSheet(f"color:{RED};")
        self._status_lbl.setText(f"Error: {msg}")

    # ── PDF export ────────────────────────────────────────────────────────

    def _on_pdf_clicked(self):
        if self._pdf_ready and self._last_pdf:
            os.startfile(self._last_pdf)
        else:
            self._export_pdf()

    def _export_pdf(self):
        if not self._last_result:
            return

        ts          = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"ont-barcoder_{ts}_fastq_ins"
        fastq_stem  = os.path.splitext(os.path.basename(self._file_path))[0]
        # strip .fastq.gz double extension if present
        if fastq_stem.endswith(".fastq"):
            fastq_stem = fastq_stem[:-6]
        pdf_name    = f"{fastq_stem}_stats.pdf"
        prog_dir    = _get_base_dir()
        auto_path   = os.path.join(prog_dir, "output", folder_name, pdf_name)

        # ── "Where to save results?" dialog — same style as other panels ──
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Output folder")
        dlg.setMinimumWidth(480)
        dlg.setStyleSheet(f"""
            QDialog {{ background-color: {GRAY_CARD}; }}
            QLabel {{ color: {TEXT_PRI}; background-color: transparent; }}
            QRadioButton {{
                color: {TEXT_PRI}; background-color: transparent;
                font-size: 15px; padding: 6px 0;
            }}
            QRadioButton::indicator {{ width: 16px; height: 16px; }}
            QPushButton {{
                border-radius: 8px; padding: 8px 20px;
                font-size: 15px; font-weight: 500;
            }}
            #dlg_ok_btn {{ background-color: {BLUE}; color: white; border: none; }}
            #dlg_ok_btn:hover {{ background-color: #0C4A82; }}
            #dlg_cancel_btn {{
                background-color: transparent; color: {BLUE};
                border: 1px solid {BLUE};
            }}
            #dlg_cancel_btn:hover {{ background-color: {BLUE_LIGHT}; }}
        """)

        vlay = QtWidgets.QVBoxLayout(dlg)
        vlay.setSpacing(16)
        vlay.setContentsMargins(24, 24, 24, 20)

        title_lbl = QtWidgets.QLabel("Where to save the results?")
        title_lbl.setStyleSheet(f"font-size:17px; font-weight:700; color:{TEXT_PRI};")
        vlay.addWidget(title_lbl)

        radio_auto = QtWidgets.QRadioButton(
            f"Automatic folder (recommended)\n  …/output/{folder_name}/{pdf_name}"
        )
        radio_auto.setChecked(True)
        radio_custom = QtWidgets.QRadioButton("Select folder manually")
        vlay.addWidget(radio_auto)
        vlay.addWidget(radio_custom)
        vlay.addSpacing(8)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_cancel.setObjectName("dlg_cancel_btn")
        btn_cancel.setFixedHeight(38)
        btn_ok = QtWidgets.QPushButton("Export PDF")
        btn_ok.setObjectName("dlg_ok_btn")
        btn_ok.setFixedHeight(38)
        btn_ok.setDefault(True)
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addSpacing(8)
        btn_row.addWidget(btn_ok)
        vlay.addLayout(btn_row)

        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return

        if radio_auto.isChecked():
            out_path = auto_path
        else:
            parent_dir = QtWidgets.QFileDialog.getExistingDirectory(
                self, "Select the output folder"
            )
            if not parent_dir:
                return
            out_path = os.path.join(parent_dir, folder_name, pdf_name)

        try:
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
        except Exception:
            pass

        try:
            from PyQt5.QtPrintSupport import QPrinter
            from PyQt5.QtGui import QTextDocument

            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(out_path)
            printer.setPageSize(QPrinter.Letter)
            printer.setOrientation(QPrinter.Portrait)
            printer.setPageMargins(15, 15, 15, 15, QPrinter.Millimeter)

            doc = QTextDocument()
            doc.setHtml(self._build_pdf_html())
            doc.print_(printer)

            self._last_pdf  = out_path
            self._pdf_ready = True
            self._pdf_btn.setText("Open PDF  ↗")

        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "PDF export", f"Could not generate PDF:\n{exc}")

    def _build_pdf_html(self) -> str:
        r = self._last_result
        n = r["n_reads"]
        fname = os.path.basename(self._file_path)
        ts = datetime.datetime.now().strftime("%Y-%m-%d  %H:%M")

        def _fmt_bp(v):
            if v >= 1_000_000_000: return f"{v/1_000_000_000:.2f} Gbp"
            if v >= 1_000_000:     return f"{v/1_000_000:.1f} Mbp"
            if v >= 1_000:         return f"{v/1_000:.1f} Kbp"
            return f"{v} bp"

        def _fmt_size(b):
            if b >= 1_073_741_824: return f"{b/1_073_741_824:.2f} GB"
            if b >= 1_048_576:     return f"{b/1_048_576:.1f} MB"
            return f"{b/1_024:.1f} KB"

        # Render charts
        imgs = {k: self._chart_to_b64(w) for k, w in (
            ("len", self._hist_len),
            ("q",   self._hist_q),
            ("gc",  self._hist_gc),
            ("sc",  self._scatter),
        )}

        css = """
        @page {
            size: Letter;
            margin: 1.5cm;
        }
        body {
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
            background: white;
            color: #2C3E50;
            margin: 0;
            padding: 0;
            font-size: 10pt;
            line-height: 1.5;
        }
        /* Header — fondo sólido + texto blanco explícito para Qt (sin gradientes ni márgenes negativos) */
        .header {
            background: #1a365d;
            color: white;
            padding: 22px 25px;
            margin-bottom: 20px;
            text-align: center;
        }
        .header h1 {
            font-size: 28pt;
            font-weight: 700;
            margin: 0 0 6px 0;
            color: white;
        }
        .header h1 strong {
            font-weight: 700;
        }
        .header .subtitle {
            font-size: 9pt;
            color: #90CDF4;
            margin-top: 8px;
        }
        .file-info {
            background: #EBF4FF;
            border-left: 4px solid #3182CE;
            padding: 12px 18px;
            margin: 20px 0;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 9pt;
        }
        /* Tarjetas de métricas */
        .metrics-grid {
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            margin: 20px 0;
            page-break-inside: avoid;
        }
        .metric-card {
            flex: 1;
            min-width: 110px;
            background: #F7FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 12px;
            text-align: center;
            page-break-inside: avoid;
        }
        .metric-card .icon {
            font-size: 22px;
            margin-bottom: 6px;
        }
        .metric-card .label {
            font-size: 8pt;
            color: #718096;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 4px;
        }
        .metric-card .value {
            font-size: 16pt;
            font-weight: 700;
            color: #2D3748;
        }
        .metric-card .unit {
            font-size: 8pt;
            color: #A0AEC0;
        }
        /* Secciones */
        .section { margin: 15px 0; }
        /* Título dentro del thead de la tabla — page-break-inside:avoid en la tabla
           es la única forma fiable de mantener título y datos juntos en QTextDocument */
        .section-title-row {
            font-size: 13pt;
            font-weight: 600;
            color: #2C5282;
            text-align: left;
            padding: 7px 12px 5px 12px;
            background: white;
            border-bottom: 2px solid #CBD5E0;
        }
        .section-title-row i { font-weight: 400; color: #718096; }
        /* Título suelto (solo sección de gráficas) */
        .section-title {
            font-size: 14pt;
            font-weight: 600;
            color: #2C5282;
            border-bottom: 2px solid #CBD5E0;
            padding-bottom: 6px;
            margin: 0 0 10px 0;
        }
        .section-title i { font-weight: 400; color: #718096; }
        /* Tablas modernas */
        .data-table {
            /* QTextDocument no respeta margin:auto (empuja la tabla a la derecha).
               Se centra con margen izquierdo EXPLÍCITO:
                 width        = ancho de la tabla
                 margin-left  = hueco a la izquierda (≈ (100-width)/2 para centrar;
                                con 80% → ~70px en Letter).
               Sube margin-left para moverla a la derecha, bájalo para acercarla
               al borde izquierdo. */
            width: 60%;
            border-collapse: collapse;
            margin: 12px 0 12px 70px;
            font-size: 9.5pt;
            page-break-inside: avoid;
        }
        .data-table th {
            background: #EDF2F7;
            color: #2D3748;
            font-weight: 600;
            padding: 7px 12px;
            text-align: left;
            border-bottom: 2px solid #CBD5E0;
        }
        .data-table td {
            padding: 5px 12px;
            border-bottom: 1px solid #E2E8F0;
        }
        .data-table td.value {
            font-weight: 600;
            color: #3182CE;
            text-align: right;
            font-family: 'Courier New', monospace;
        }
        /* Notas bajo las gráficas */
        .chart-note {
            font-size: 8pt;
            color: #A0AEC0;
            text-align: center;
            margin-top: 4px;
        }
        /* Footer */
        .footer {
            margin-top: 35px;
            padding-top: 15px;
            text-align: center;
            font-size: 8pt;
            color: #A0AEC0;
            border-top: 1px solid #E2E8F0;
        }
        /* Badges de calidad */
        .badge-good { background: #C6F6D5; color: #22543D; padding: 2px 8px; border-radius: 20px; font-size: 8pt; font-weight: 600; }
        .badge-warning { background: #FEEBC8; color: #7B341E; padding: 2px 8px; border-radius: 20px; font-size: 8pt; font-weight: 600; }
        .badge-info { background: #BEE3F8; color: #1A365D; padding: 2px 8px; border-radius: 20px; font-size: 8pt; font-weight: 600; }
        """

        # Calcular indicadores de calidad
        q_mean = r['mean_q']
        q_badge = 'badge-good' if q_mean >= 20 else ('badge-warning' if q_mean >= 15 else 'badge-info')
        q_badge_text = 'Excellent' if q_mean >= 20 else ('Good' if q_mean >= 15 else 'Fair')

        # Tarjetas de métricas
        metrics = f"""
        <div class="metrics-grid">
            <div class="metric-card">
                <div class="icon">📖</div>
                <div class="label">Total Reads</div>
                <div class="value">{n:,}</div>
            </div>
            <div class="metric-card">
                <div class="icon">🧬</div>
                <div class="label">Total Bases</div>
                <div class="value">{_fmt_bp(r['total_bases'])}</div>
            </div>
            <div class="metric-card">
                <div class="icon">📏</div>
                <div class="label">N50</div>
                <div class="value">{r['n50']:,}</div>
                <div class="unit">bp</div>
            </div>
            <div class="metric-card">
                <div class="icon">⭐</div>
                <div class="label">Mean Quality</div>
                <div class="value">Q{r['mean_q']:.1f}</div>
                <div class="unit"><span class="{q_badge}">{q_badge_text}</span></div>
            </div>
            <div class="metric-card">
                <div class="icon">🧪</div>
                <div class="label">GC Content</div>
                <div class="value">{r['mean_gc']:.1f}%</div>
            </div>
        </div>
        """

        # Tabla de longitud — título dentro de thead para que page-break-inside:avoid
        # en la tabla mantenga título y filas juntos (QTextDocument no respeta el
        # page-break en <div>, pero sí en elementos <table>)
        length_table = f"""
        <table class="data-table" style="page-break-before: always; margin-top: 0;">
            <thead>
                <tr><th colspan="3" class="section-title-row">📏 Length Distribution <i>— read length statistics</i></th></tr>
                <tr><th>Metric</th><th>Value</th><th></th></tr>
            </thead>
            <tbody>
                <tr><td>Minimum</td><td class="value">{r['min_len']:,} bp</td><td></td></tr>
                <tr><td>Maximum</td><td class="value">{r['max_len']:,} bp</td><td></td></tr>
                <tr><td>Mean</td><td class="value">{r['mean_len']:,.0f} bp</td><td></td></tr>
                <tr><td>Median</td><td class="value">{r['median_len']:,.0f} bp</td><td></td></tr>
                <tr><td>N50</td><td class="value">{r['n50']:,} bp</td><td><span class="badge-info">Assembly standard</span></td></tr>
                <tr><td>&lt; 500 bp</td><td class="value">{r['len_lt500']:,}</td><td>({r['len_lt500']/n*100:.1f}%)</td></tr>
                <tr><td>500–2000 bp</td><td class="value">{r['len_500_2k']:,}</td><td>({r['len_500_2k']/n*100:.1f}%)</td></tr>
                <tr><td>&gt; 2000 bp</td><td class="value">{r['len_gt2k']:,}</td><td>({r['len_gt2k']/n*100:.1f}%)</td></tr>
            </tbody>
        </table>
        """

        # Tabla de calidad
        quality_table = f"""
        <table class="data-table">
            <thead>
                <tr><th colspan="3" class="section-title-row">⭐ Base Quality <i>— Phred score distribution</i></th></tr>
                <tr><th>Metric</th><th>Value</th><th>Quality assessment</th></tr>
            </thead>
            <tbody>
                <tr><td>Mean Q-score</td><td class="value">Q{r['mean_q']:.1f}</td><td><span class="{q_badge}">{q_badge_text}</span></td></tr>
                <tr><td>Reads ≥ Q10</td><td class="value">{r['q10_pct']:.1f}%</td><td></td></tr>
                <tr><td>Reads ≥ Q15</td><td class="value">{r['q15_pct']:.1f}%</td><td></td></tr>
                <tr><td>Reads ≥ Q20</td><td class="value">{r['q20_pct']:.1f}%</td><td><span class="badge-good">High confidence</span></td></tr>
            </tbody>
        </table>
        """

        # Tabla de GC
        gc_table = f"""
        <table class="data-table">
            <thead>
                <tr><th colspan="3" class="section-title-row">🧬 GC Content <i>— nucleotide composition</i></th></tr>
                <tr><th>Metric</th><th>Value</th><th>Expected range</th></tr>
            </thead>
            <tbody>
                <tr><td>Mean GC content</td><td class="value">{r['mean_gc']:.1f}%</td><td>40–60% for most organisms</td></tr>
                <tr><td>Standard deviation</td><td class="value">±{r['std_gc']:.1f}%</td><td>Lower = more homogeneous</td></tr>
            </tbody>
        </table>
        """

        # Las dos gráficas de cada hoja van en UNA sola <table> y la celda de cada
        # imagen lleva line-height:100% (ver _img_row). Sin eso, el line-height:1.5
        # del body inflaba cada imagen ~1.5× y solo entraba una gráfica por hoja.
        # Resultado: dos gráficas por hoja.
        _cw = self._PDF_CHART_W   # PNG ya embebido a este tamaño exacto
        _ch = self._PDF_CHART_H

        def _title_row(title, top_border=False):
            bt = "border-top:1px solid #FBFBFB;" if top_border else ""
            return (
                f'<tr><td style="padding:5px 8px; font-size:10pt; font-weight:600;'
                f' color:#4A5568; border-bottom:1px solid #E2E8F0;{bt}">{title}</td></tr>'
            )

        # ▼▼ Espacio (px) DEBAJO de cada gráfica, antes del título de la siguiente.
        #    Sube/baja este número para ajustar ese hueco. Ojo: si lo subes mucho,
        #    la 2ª gráfica podría no caber en la hoja.
        _gap_below_chart = 40

        def _img_row(b64):
            # line-height:100% es clave: QTextDocument aplica el line-height (1.5
            # heredado del body) de forma MULTIPLICATIVA a la línea que contiene la
            # imagen, así que cada gráfica de 280px ocupaba ~420px y sobraban ~140px
            # de hueco debajo (lo que empujaba la 2ª gráfica a otra hoja). Con 100%
            # la línea mide exactamente la altura de la imagen.
            return (
                f'<tr><td style="padding:4px 8px {_gap_below_chart}px; line-height:100%;">'
                f'<img src="data:image/png;base64,{b64}" width="{_cw}" height="{_ch}"/>'
                f'</td></tr>'
            )

        _header_row = (
            '<tr><td style="padding:2px 8px 8px; font-size:14pt; font-weight:600;'
            ' color:#2C5282; border-bottom:2px solid #CBD5E0;">'
            '📈 Visual Summary <i style="font-weight:400; color:#718096;">'
            '— distribution plots</i></td></tr>'
        )
        _page_open = ('<table width="100%" cellspacing="0" cellpadding="0"'
                      ' style="page-break-before:always; border:1px solid #E2E8F0;'
                      ' border-collapse:collapse; margin:0;">')

        charts = f"""
    {_page_open}
        {_header_row}
        {_title_row("📊 Read length distribution")}
        {_img_row(imgs['len'])}
        {_title_row("📈 Mean Q-score per read", top_border=True)}
        {_img_row(imgs['q'])}
</table>
    {_page_open}
        {_title_row("🧬 GC content per read")}
        {_img_row(imgs['gc'])}
        {_title_row("📐 Length vs. Quality", top_border=True)}
        {_img_row(imgs['sc'])}
</table>
        """

        return f"""<!DOCTYPE html>
    <html><head><meta charset="utf-8"><style>{css}</style></head>
    <body>

    <div class="header">
        <h1><strong>FASTQ</strong> Inspector Report</h1>
        <div class="subtitle">Generated: {ts}</div>
    </div>

    <div class="file-info">
        📁 <strong>Input file:</strong> {fname}<br>
        💾 <strong>File size:</strong> {_fmt_size(self._file_size)}
    </div>

    {metrics}
    {length_table}
    {quality_table}
    {gc_table}
    {charts}

    </body></html>"""


# ═══════════════════════════════════════════════════════════════════════════
# FASTA TOOLS
# ═══════════════════════════════════════════════════════════════════════════


class _FastaToolsWorker(QtCore.QThread):
    progress = QtCore.pyqtSignal(str)   # short status shown in the label
    log_line = QtCore.pyqtSignal(str)   # detailed per-file result appended to the log
    finished = QtCore.pyqtSignal(list)
    error    = QtCore.pyqtSignal(str)

    def __init__(self, files: list, operation: str, params: dict, out_dir: str, parent=None):
        super().__init__(parent)
        self._files     = files
        self._operation = operation
        self._params    = params
        self._out_dir   = out_dir
        self._stop      = False

    def stop(self):
        self._stop = True

    @staticmethod
    def _read_fasta(path):
        import gzip
        opener = gzip.open if path.lower().endswith(".gz") else open
        records = []
        header = None
        seq_parts = []
        with opener(path, "rt", encoding="utf-8", errors="replace") as fh:
            for line in fh:
                line = line.rstrip("\r\n")
                if line.startswith(">"):
                    if header is not None:
                        records.append((header, "".join(seq_parts)))
                    header = line[1:]
                    seq_parts = []
                elif header is not None:
                    seq_parts.append(line)
            if header is not None:
                records.append((header, "".join(seq_parts)))
        return records

    @staticmethod
    def _write_fasta(records, path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            for hdr, seq in records:
                fh.write(f">{hdr}\n{seq}\n")

    @staticmethod
    def _stem(path):
        name = os.path.basename(path)
        for ext in (".fasta.gz", ".fas.gz", ".fa.gz", ".fna.gz",
                    ".fasta", ".fas", ".fa", ".fna"):
            if name.lower().endswith(ext):
                return name[: -len(ext)]
        return os.path.splitext(name)[0]

    def _iter_inputs(self):
        """Yield (display_name, stem, records) per input to process.
        In merge mode, concatenates all files and yields a single entry.
        When two files share the same stem, prefixes with the parent folder name."""
        if self._params.get("merge") and len(self._files) > 1:
            self.progress.emit("Merging files…")
            all_records = []
            for path in self._files:
                if self._stop:
                    return
                all_records.extend(self._read_fasta(path))
            yield f"Merged {len(self._files)} files", "merged", all_records
        else:
            stems = [self._stem(p) for p in self._files]
            seen_stems = {s for s in stems if stems.count(s) > 1}
            used_stems: set = set()
            for path, stem in zip(self._files, stems):
                if self._stop:
                    return
                name = os.path.basename(path)
                self.progress.emit(f"Processing {name}…")
                if stem in seen_stems:
                    parent = os.path.basename(os.path.dirname(os.path.abspath(path)))
                    stem = f"{parent}_{stem}"
                # Guarantee a unique output stem even when the parent folder also
                # matches, so a later file never silently overwrites an earlier one.
                base_stem = stem
                counter = 2
                while stem in used_stems:
                    stem = f"{base_stem}_{counter}"
                    counter += 1
                used_stems.add(stem)
                yield name, stem, self._read_fasta(path)

    @staticmethod
    def _cell_to_str(v) -> str:
        """Convert an Excel cell value to a string for ID matching.
        openpyxl returns integer-valued cells as float (e.g. 1234.0); turn those
        into "1234" so they match a "1234" FASTA id instead of becoming "1234_0"
        after the dot is normalized to an underscore."""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    @staticmethod
    def _normalize_field(s: str) -> str:
        import unicodedata, re as _re
        s = str(s).strip()
        s = unicodedata.normalize("NFD", s)
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        s = _re.sub(r"[ .,]+", "_", s)
        s = _re.sub(r"_+", "_", s)
        return s.strip("_")

    def _run_unique(self):
        outputs = []
        for display_name, stem, records in self._iter_inputs():
            if self._stop:
                break
            seen = set()
            unique = []
            for hdr, seq in records:
                key = seq.upper()
                if key not in seen:
                    seen.add(key)
                    unique.append((hdr, seq))
            removed = len(records) - len(unique)
            out_path = os.path.join(self._out_dir, f"{stem}_unique.fasta")
            self._write_fasta(unique, out_path)
            self.log_line.emit(
                f"{display_name}\n"
                f"  Total sequences in :  {len(records)}\n"
                f"  Unique sequences   :  {len(unique)}\n"
                f"  Duplicates removed :  {removed}\n"
            )
            outputs.append(out_path)
        return outputs

    def _run_identical(self):
        outputs = []
        for display_name, stem, records in self._iter_inputs():
            if self._stop:
                break
            counts = Counter(seq.upper() for _, seq in records)
            identical = [(hdr, seq) for hdr, seq in records if counts[seq.upper()] > 1]
            dup_groups = sum(1 for c in counts.values() if c > 1)

            # FASTA output
            out_path = os.path.join(self._out_dir, f"{stem}_identical.fasta")
            self._write_fasta(identical, out_path)

            # Excel: group IDs by sequence
            seq_to_ids: dict = {}
            for hdr, seq in records:
                key = seq.upper()
                if counts[key] > 1:
                    seq_to_ids.setdefault(key, []).append(hdr)

            # ── Excel: two sheets (IDs + Groups) ────────────────────────────
            xlsx_path = os.path.join(self._out_dir, f"{stem}_duplicate_groups.xlsx")
            wb = xlsxwriter.Workbook(xlsx_path)
            hf = wb.add_format({"bold": True, "bg_color": "#4F81BD",
                                "font_color": "#FFFFFF", "border": 1})
            cf = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
            nf = wb.add_format({"border": 1, "align": "center", "valign": "top"})

            ws_i = wb.add_worksheet("IDs")
            ws_i.activate()
            ws_g = wb.add_worksheet("Groups")

            for col, h in enumerate(["Group", "ID"]):
                ws_i.write(0, col, h, hf)
            ws_i.set_column(0, 0, 8)
            ws_i.set_column(1, 1, 30)

            for col, h in enumerate(["Group", "Count", "Sequence"]):
                ws_g.write(0, col, h, hf)
            ws_g.set_column(0, 1, 8)
            ws_g.set_column(2, 2, 50)

            id_row = 1
            for g_idx, (seq_key, ids) in enumerate(seq_to_ids.items(), start=1):
                ws_g.write(g_idx, 0, g_idx, nf)
                ws_g.write(g_idx, 1, len(ids), nf)
                # Excel caps cell text at 32767 chars; truncate longer sequences
                # so xlsxwriter does not raise/drop the cell silently.
                if len(seq_key) > 32767:
                    ws_g.write(g_idx, 2, seq_key[:32764] + "...", cf)
                else:
                    ws_g.write(g_idx, 2, seq_key, cf)
                for id_val in ids:
                    ws_i.write(id_row, 0, g_idx, nf)
                    ws_i.write(id_row, 1, id_val, cf)
                    id_row += 1
            wb.close()
            outputs.append(xlsx_path)

            self.log_line.emit(
                f"{display_name}\n"
                f"  Total sequences in       :  {len(records)}\n"
                f"  Identical sequences out  :  {len(identical)}\n"
                f"  Distinct duplicated seqs :  {dup_groups}\n"
                f"  Excel groups file        :  {os.path.basename(xlsx_path)}\n"
            )
            outputs.append(out_path)
        return outputs

    def _run_grep(self):
        import re
        pattern_str  = self._params.get("pattern", "")
        pattern_file = self._params.get("pattern_file", "")
        use_regex    = self._params.get("use_regex", False)

        patterns = []
        if pattern_file and os.path.isfile(pattern_file):
            with open(pattern_file, encoding="utf-8", errors="replace") as f:
                patterns = [ln.strip() for ln in f if ln.strip()]
        elif pattern_str:
            patterns = [pattern_str]

        if not patterns:
            raise ValueError("No pattern specified for grep operation.")

        mode_str = "regex" if use_regex else "plain text"
        pat_summary = patterns[0] if len(patterns) == 1 else f"{len(patterns)} patterns from file"

        if use_regex:
            compiled = [re.compile(p) for p in patterns]
            def matches(hdr):
                return any(rx.search(hdr) for rx in compiled)
        else:
            def matches(hdr):
                return any(p in hdr for p in patterns)

        outputs = []
        for display_name, stem, records in self._iter_inputs():
            if self._stop:
                break
            matched = [(hdr, seq) for hdr, seq in records if matches(hdr)]
            unmatched = len(records) - len(matched)
            out_path = os.path.join(self._out_dir, f"{stem}_grep.fasta")
            self._write_fasta(matched, out_path)
            self.log_line.emit(
                f"{display_name}\n"
                f"  Pattern ({mode_str})  :  {pat_summary}\n"
                f"  Total sequences in  :  {len(records)}\n"
                f"  Matched             :  {len(matched)}\n"
                f"  Not matched         :  {unmatched}\n"
            )
            outputs.append(out_path)
        return outputs

    def _run_append(self):
        import openpyxl
        excel_path = self._params.get("excel_file", "")
        separator  = self._params.get("separator", "|")

        if not excel_path or not os.path.isfile(excel_path):
            raise ValueError("No valid Excel file specified.")

        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active
        lookup: dict = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            # Key is normalized so dots, commas and spaces become underscores —
            # this makes Excel IDs like "Col_1.1" match FASTA IDs like "Col_1_1"
            row_id = self._normalize_field(self._cell_to_str(row[0]))
            fields = [
                nv for v in row[1:]
                if v is not None and self._cell_to_str(v).strip()
                for nv in (self._normalize_field(self._cell_to_str(v)),)
                if any(c.isalnum() for c in nv)
            ]
            lookup[row_id] = fields

        excel_name = os.path.basename(excel_path)
        n_excel_ids = len(lookup)

        outputs = []
        for display_name, stem, records in self._iter_inputs():
            if self._stop:
                break
            updated_list = []
            n_updated = 0
            n_not_found = 0
            for hdr, seq in records:
                hdr_parts = hdr.split()
                seq_id = self._normalize_field(hdr_parts[0]) if hdr_parts else ""
                fields = lookup.get(seq_id, []) if seq_id else []
                if fields:
                    new_hdr = hdr + separator + separator.join(fields)
                    n_updated += 1
                else:
                    new_hdr = hdr
                    n_not_found += 1
                updated_list.append((new_hdr, seq))
            out_path = os.path.join(self._out_dir, f"{stem}_append.fasta")
            self._write_fasta(updated_list, out_path)
            self.log_line.emit(
                f"{display_name}\n"
                f"  Excel file          :  {excel_name}  ({n_excel_ids} IDs)\n"
                f"  Total sequences in  :  {len(records)}\n"
                f"  IDs updated         :  {n_updated}\n"
                f"  IDs not in Excel    :  {n_not_found}\n"
            )
            outputs.append(out_path)
        return outputs

    def _run_reformat(self):
        mode      = self._params.get("mode", "linearize")   # "linearize" | "wrap"
        wrap_cols = int(self._params.get("wrap_cols", 80))

        if mode == "linearize":
            suffix = "_linearized"
            mode_desc = "multi-line → single line"
        else:
            def _wrap(seq, w=wrap_cols):
                return "\n".join(seq[i: i + w] for i in range(0, len(seq), w)) if seq else ""
            suffix = "_wrapped"
            mode_desc = f"single line → multi-line  ({wrap_cols} nt/row)"

        outputs = []
        for display_name, stem, records in self._iter_inputs():
            if self._stop:
                break
            reformatted = records if mode == "linearize" else [(hdr, _wrap(seq)) for hdr, seq in records]
            out_path = os.path.join(self._out_dir, f"{stem}{suffix}.fasta")
            self._write_fasta(reformatted, out_path)
            self.log_line.emit(
                f"{display_name}\n"
                f"  Mode                :  {mode_desc}\n"
                f"  Total sequences     :  {len(records)}\n"
            )
            outputs.append(out_path)
        return outputs

    def _run_sort(self):
        separator = self._params.get("separator", "|")   # "" → sort whole header
        levels    = self._params.get("levels", [{"field": 1, "order": "asc"}])

        sep_desc = f'"{separator}"' if separator else "none (whole header)"
        levels_desc = "  |  ".join(
            f"Field {lv['field']} {'↑' if lv['order'] == 'asc' else '↓'}"
            for lv in levels
        )

        outputs = []
        for display_name, stem, records in self._iter_inputs():
            if self._stop:
                break
            working = list(records)
            # Stable multi-key sort: apply levels from last to first.
            for level in reversed(levels):
                idx     = level["field"] - 1
                reverse = level["order"] == "desc"

                def _key(rec, i=idx, sep=separator):
                    parts = rec[0].split(sep) if sep else [rec[0]]
                    val   = parts[i].strip() if i < len(parts) else ""
                    # Natural sort: split into (text, int) chunks so that
                    # "Sample_10" > "Sample_2" instead of "Sample_10" < "Sample_2".
                    return tuple(
                        int(c) if c.isdigit() else c.lower()
                        for c in re.split(r"(\d+)", val)
                    )

                working = sorted(working, key=_key, reverse=reverse)

            out_path = os.path.join(self._out_dir, f"{stem}_sorted.fasta")
            self._write_fasta(working, out_path)
            self.log_line.emit(
                f"{display_name}\n"
                f"  Total sequences  :  {len(records)}\n"
                f"  Separator        :  {sep_desc}\n"
                f"  Sort levels      :  {levels_desc}\n"
            )
            outputs.append(out_path)
        return outputs

    def _run_filter_fields(self):
        separator = self._params.get("separator", "|")
        criteria  = self._params.get("criteria", [])
        if not criteria:
            raise ValueError("No filter criteria specified.")

        parsed = []
        for c in criteria:
            idx = c.get("field", 1) - 1
            op  = c.get("op", "=")
            raw = c.get("value", "").strip()
            if op in ("contains", "not contains"):
                val     = raw.lower()
                numeric = False
            else:
                try:
                    val     = float(raw)
                    numeric = True
                except ValueError:
                    val     = raw.lower()
                    numeric = False
            parsed.append((idx, op, val, numeric))

        def _field_value(header, idx):
            parts = header.split(separator) if separator else [header]
            if idx >= len(parts):
                return None
            raw = parts[idx].strip()
            if "=" in raw:
                raw = raw.split("=", 1)[1].strip()
            return raw

        def _passes(header):
            for idx, op, val, numeric in parsed:
                raw_v = _field_value(header, idx)
                if raw_v is None:
                    return False
                if numeric:
                    try:
                        fv = float(raw_v)
                    except ValueError:
                        return False
                    if op == ">="           and not (fv >= val): return False
                    if op == "<="           and not (fv <= val): return False
                    if op == ">"            and not (fv >  val): return False
                    if op == "<"            and not (fv <  val): return False
                    if op == "="            and not (fv == val): return False
                    if op == "!="           and not (fv != val): return False
                else:
                    sv = raw_v.lower()
                    if op == "="            and sv != val:         return False
                    if op == "!="           and sv == val:         return False
                    if op == "contains"     and val not in sv:     return False
                    if op == "not contains" and val in sv:         return False
            return True

        sep_desc  = f'"{separator}"' if separator else "none (whole header)"
        crit_desc = "  AND  ".join(
            f"field {i + 1} {o} {v}" for i, o, v, _ in parsed
        )
        outputs = []
        for display_name, stem, records in self._iter_inputs():
            if self._stop:
                break
            matched   = [(h, s) for h, s in records if _passes(h)]
            unmatched = len(records) - len(matched)
            out_path  = os.path.join(self._out_dir, f"{stem}_filtered.fasta")
            self._write_fasta(matched, out_path)
            self.log_line.emit(
                f"{display_name}\n"
                f"  Separator        :  {sep_desc}\n"
                f"  Criteria (AND)   :  {crit_desc}\n"
                f"  Total sequences  :  {len(records)}\n"
                f"  Passed           :  {len(matched)}\n"
                f"  Excluded         :  {unmatched}\n"
            )
            outputs.append(out_path)
        return outputs

    def run(self):
        try:
            os.makedirs(self._out_dir, exist_ok=True)
            if self._operation == "unique":
                outputs = self._run_unique()
            elif self._operation == "identical":
                outputs = self._run_identical()
            elif self._operation == "grep":
                outputs = self._run_grep()
            elif self._operation == "filter_fields":
                outputs = self._run_filter_fields()
            elif self._operation == "append":
                outputs = self._run_append()
            elif self._operation == "reformat":
                outputs = self._run_reformat()
            elif self._operation == "sort":
                outputs = self._run_sort()
            else:
                raise ValueError(f"Unknown operation: {self._operation}")
            if not self._stop:
                self.finished.emit(outputs)
        except Exception as exc:
            self.error.emit(str(exc))


class _DragDropLineEdit(QtWidgets.QLineEdit):
    """QLineEdit that accepts a single file dragged onto it."""

    _DRAG_STYLE = (
        f"QLineEdit {{ border: 2px solid {BLUE_MID}; background-color: {BLUE_LIGHT};"
        f" border-radius: 4px; }}"
    )

    def __init__(self, accepted_extensions=None, parent=None):
        super().__init__(parent)
        self._accepted_ext = [e.lower() for e in (accepted_extensions or [])]
        self.setAcceptDrops(True)

    def _is_valid(self, url):
        if not url.isLocalFile():
            return False
        if self._accepted_ext:
            return any(url.toLocalFile().lower().endswith(e) for e in self._accepted_ext)
        return True

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and any(self._is_valid(u) for u in event.mimeData().urls()):
            event.acceptProposedAction()
            self.setStyleSheet(self._DRAG_STYLE)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls() and any(self._is_valid(u) for u in event.mimeData().urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("")
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        self.setStyleSheet("")
        for url in event.mimeData().urls():
            if self._is_valid(url):
                self.setText(url.toLocalFile())
                event.acceptProposedAction()
                return
        event.ignore()


class FastaToolsPanel(QtWidgets.QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._sort_field_count: int = 99

        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── Scroll area ──
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        inner = QtWidgets.QWidget()
        self._layout = QtWidgets.QVBoxLayout(inner)
        self._layout.setContentsMargins(20, 20, 20, 20)
        self._layout.setSpacing(16)
        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        # ── Title ──
        self._layout.addWidget(make_label("FASTA Tools", size=19, bold=True))
        desc = make_label(
            "Drag one or multiple FASTA files (.fa, .fas, .fasta) and choose an operation to run on them.",
            color=TEXT_SEC,
        )
        desc.setWordWrap(True)
        self._layout.addWidget(desc)

        # ── Drop zone (multiple files) ──
        self._drop = MultiDropZone()
        self._drop.filesDropped.connect(self._on_files)
        self._layout.addWidget(self._drop)

        # ── Merge option ──
        self._merge_chk = QtWidgets.QCheckBox("Merge files before processing")
        self._merge_chk.setToolTip(
            "Checked: all files are concatenated into a single dataset\n"
            "and the operation runs on the combined sequences.\n"
            "Output is saved as a single merged file.\n\n"
            "Unchecked: each file is processed individually\n"
            "and a separate output file is generated for each one."
        )
        self._layout.addWidget(self._merge_chk)

        # ── Operation selector — each radio + its params as an inline block ──
        self._layout.addWidget(self._make_section_lbl("Operation"))

        ops_box = QtWidgets.QWidget()
        ops_layout = QtWidgets.QVBoxLayout(ops_box)
        ops_layout.setContentsMargins(0, 0, 0, 0)
        ops_layout.setSpacing(2)

        self._radio_unique    = QtWidgets.QRadioButton("Extract unique sequences")
        self._radio_identical = QtWidgets.QRadioButton("Extract identical sequences (duplicates)")
        self._radio_identical.setToolTip(
            "Extracts sequences that appear more than once.\n"
            "Generates two output files:\n"
            "  • <name>_identical.fasta — duplicated sequences\n"
            "  • <name>_duplicate_groups.xlsx — IDs grouped by shared sequence\n"
            "    · Sheet 'IDs': list of IDs per group\n"
            "    · Sheet 'Groups': group number, ID count and sequence"
        )
        self._radio_grep          = QtWidgets.QRadioButton("Extract sequences by pattern (grep)")
        self._radio_filter_fields = QtWidgets.QRadioButton("Filter sequences by header fields")
        self._radio_append        = QtWidgets.QRadioButton("Append info to headers from .xlsx file")
        self._radio_reformat      = QtWidgets.QRadioButton("Reformat sequence lines")
        self._radio_sort          = QtWidgets.QRadioButton("Sort sequences")

        self._op_group = QtWidgets.QButtonGroup(self)
        for rb in (self._radio_unique, self._radio_identical,
                   self._radio_grep, self._radio_filter_fields,
                   self._radio_append, self._radio_reformat, self._radio_sort):
            self._op_group.addButton(rb)

        self._radio_unique.setChecked(True)
        self._op_group.buttonClicked.connect(self._on_operation_changed)

        # Operations without params
        ops_layout.addWidget(self._radio_unique)
        ops_layout.addWidget(self._radio_identical)

        # ── Grep options ──
        ops_layout.addSpacing(4)
        ops_layout.addWidget(self._radio_grep)
        self._grep_widget = QtWidgets.QWidget()
        gl = QtWidgets.QVBoxLayout(self._grep_widget)
        gl.setContentsMargins(20, 4, 0, 4)
        gl.setSpacing(8)

        self._grep_use_file_chk = QtWidgets.QCheckBox("Load patterns from file")
        self._grep_use_file_chk.stateChanged.connect(self._on_grep_source_changed)
        gl.addWidget(self._grep_use_file_chk)

        self._grep_pattern_row = QtWidgets.QWidget()
        gpr = QtWidgets.QHBoxLayout(self._grep_pattern_row)
        gpr.setContentsMargins(0, 0, 0, 0)
        gpr.setSpacing(8)
        gpr.addWidget(make_label("Pattern:", color=TEXT_SEC))
        self._grep_pattern_edit = QtWidgets.QLineEdit()
        self._grep_pattern_edit.setPlaceholderText("Enter search pattern…")
        gpr.addWidget(self._grep_pattern_edit, 1)
        gl.addWidget(self._grep_pattern_row)

        self._grep_file_row = QtWidgets.QWidget()
        gfr = QtWidgets.QHBoxLayout(self._grep_file_row)
        gfr.setContentsMargins(0, 0, 0, 0)
        gfr.setSpacing(8)
        lbl_pf = make_label("Patterns file:", color=TEXT_SEC)
        lbl_pf.setToolTip(
            "Plain text file — one pattern per line.\n"
            "Empty lines are ignored.\n"
            "Each pattern is matched against the full FASTA header.\n"
            "Enable 'Regex' to treat each line as a regular expression."
        )
        gfr.addWidget(lbl_pf)
        self._grep_file_edit = _DragDropLineEdit(accepted_extensions=[".txt"])
        self._grep_file_edit.setReadOnly(True)
        self._grep_file_edit.setPlaceholderText("No file selected… (or drag & drop)")
        self._grep_file_edit.setToolTip(
            "Plain text file — one pattern per line.\n"
            "Empty lines are ignored.\n"
            "Each pattern is matched against the full FASTA header.\n"
            "Enable 'Regex' to treat each line as a regular expression."
        )
        gfr.addWidget(self._grep_file_edit, 1)
        self._grep_file_clear = self._make_clear_btn()
        self._grep_file_clear.setVisible(False)
        self._grep_file_clear.clicked.connect(lambda: (
            self._grep_file_edit.clear(),
            self._grep_file_clear.setVisible(False),
        ))
        self._grep_file_edit.textChanged.connect(
            lambda t: self._grep_file_clear.setVisible(bool(t)))
        gfr.addWidget(self._grep_file_clear)
        self._grep_file_btn = QtWidgets.QPushButton("Browse…")
        self._grep_file_btn.setObjectName("secondary_btn")
        self._grep_file_btn.setFixedWidth(120)
        self._grep_file_btn.clicked.connect(self._browse_grep_file)
        gfr.addWidget(self._grep_file_btn)
        self._grep_file_row.hide()
        gl.addWidget(self._grep_file_row)

        _grep_file_note = make_label(
            "Plain text file · one pattern per line · empty lines ignored",
            size=14, color=TEXT_HINT,
        )
        gl.addWidget(_grep_file_note)
        self._grep_file_note = _grep_file_note
        _grep_file_note.hide()

        self._grep_regex_chk = QtWidgets.QCheckBox("Use regex (regular expression)")
        gl.addWidget(self._grep_regex_chk)
        self._grep_widget.hide()
        ops_layout.addWidget(self._grep_widget)

        # ── Filter by header fields options ──
        ops_layout.addSpacing(4)
        ops_layout.addWidget(self._radio_filter_fields)
        self._filter_widget = QtWidgets.QWidget()
        ffw = QtWidgets.QVBoxLayout(self._filter_widget)
        ffw.setContentsMargins(20, 4, 0, 4)
        ffw.setSpacing(10)

        ff_note = make_label(
            "Split each header by a separator and filter by field position. "
            "All criteria must be satisfied (AND). "
            "Numeric operators: >=  <=  >  <  =  !=   ·   Text operators: =  !=  contains  not contains",
            color=TEXT_SEC, size=15,
        )
        ff_note.setWordWrap(True)
        ffw.addWidget(ff_note)

        ff_sep_row = QtWidgets.QHBoxLayout()
        ff_sep_row.setSpacing(12)
        ff_sep_row.addWidget(make_label("Field separator:", color=TEXT_SEC))
        self._ff_sep_group = QtWidgets.QButtonGroup(self)
        self._ff_sep_pipe   = QtWidgets.QRadioButton('"|"  pipe')
        self._ff_sep_semi   = QtWidgets.QRadioButton('";"  semicolon')
        self._ff_sep_custom = QtWidgets.QRadioButton("Custom:")
        self._ff_sep_pipe.setChecked(True)
        for rb in (self._ff_sep_pipe, self._ff_sep_semi, self._ff_sep_custom):
            self._ff_sep_group.addButton(rb)
            ff_sep_row.addWidget(rb)
        self._ff_sep_custom_edit = QtWidgets.QLineEdit()
        self._ff_sep_custom_edit.setFixedWidth(64)
        self._ff_sep_custom_edit.setPlaceholderText("e.g. _")
        self._ff_sep_custom_edit.setMaxLength(5)
        self._ff_sep_custom_edit.hide()
        ff_sep_row.addWidget(self._ff_sep_custom_edit)
        ff_sep_row.addStretch()
        ffw.addLayout(ff_sep_row)
        self._ff_sep_group.buttonClicked.connect(self._on_ff_sep_changed)
        self._ff_sep_custom_edit.textChanged.connect(self._update_ff_preview)

        self._ff_preview_frame = QtWidgets.QFrame()
        self._ff_preview_frame.setObjectName("ff_preview_card")
        self._ff_preview_frame.setStyleSheet(f"""
            QFrame#ff_preview_card {{
                background: {GRAY_BG};
                border: 1px solid {GRAY_LINE};
                border-radius: 8px;
            }}
        """)
        ff_pfl = QtWidgets.QVBoxLayout(self._ff_preview_frame)
        ff_pfl.setContentsMargins(12, 8, 12, 10)
        ff_pfl.setSpacing(6)
        ff_pfl.addWidget(make_label("Field preview  (first loaded sequence):", size=15, color=TEXT_SEC))
        self._ff_preview_lbl = QtWidgets.QLabel("—  no files loaded")
        self._ff_preview_lbl.setStyleSheet(
            f"font-family:'Courier New',Consolas,monospace; font-size:14px; color:{TEXT_PRI};"
            f" background:transparent;"
        )
        self._ff_preview_lbl.setWordWrap(True)
        ff_pfl.addWidget(self._ff_preview_lbl)
        ffw.addWidget(self._ff_preview_frame)

        ffw.addWidget(make_label("Criteria (all must match):", color=TEXT_SEC))

        self._filter_criteria_container = QtWidgets.QWidget()
        self._filter_criteria_layout = QtWidgets.QVBoxLayout(self._filter_criteria_container)
        self._filter_criteria_layout.setContentsMargins(0, 0, 0, 0)
        self._filter_criteria_layout.setSpacing(6)
        ffw.addWidget(self._filter_criteria_container)

        self._filter_criterion_rows: list = []
        self._ff_field_count: int = 99
        self._add_filter_criterion()

        add_crit_btn = QtWidgets.QPushButton("+ Add criterion")
        add_crit_btn.setObjectName("secondary_btn")
        add_crit_btn.setFixedWidth(180)
        add_crit_btn.clicked.connect(self._add_filter_criterion)
        ffw.addWidget(add_crit_btn)

        self._filter_widget.hide()
        ops_layout.addWidget(self._filter_widget)

        # ── Append options ──
        ops_layout.addSpacing(4)
        ops_layout.addWidget(self._radio_append)
        self._append_widget = QtWidgets.QWidget()
        al = QtWidgets.QVBoxLayout(self._append_widget)
        al.setContentsMargins(20, 4, 0, 4)
        al.setSpacing(8)

        note = make_label(
            "Excel format: column 1 = sequence ID (same as FASTA header, without '>'), "
            "columns 2+ = fields to append. Cell values are normalized: "
            "leading/trailing spaces removed, spaces/dots/commas → '_', accents removed.",
            color=TEXT_SEC, size=15,
        )
        note.setWordWrap(True)
        al.addWidget(note)

        excel_row = QtWidgets.QHBoxLayout()
        excel_row.setSpacing(8)
        excel_row.addWidget(make_label("Excel file:", color=TEXT_SEC))
        self._excel_edit = _DragDropLineEdit(accepted_extensions=[".xlsx", ".xls"])
        self._excel_edit.setReadOnly(True)
        self._excel_edit.setPlaceholderText("No file selected… (or drag & drop)")
        excel_row.addWidget(self._excel_edit, 1)
        self._excel_clear = self._make_clear_btn()
        self._excel_clear.setVisible(False)
        self._excel_clear.clicked.connect(lambda: (
            self._excel_edit.clear(),
            self._excel_clear.setVisible(False),
        ))
        self._excel_edit.textChanged.connect(
            lambda t: self._excel_clear.setVisible(bool(t)))
        excel_row.addWidget(self._excel_clear)
        self._excel_btn = QtWidgets.QPushButton("Browse…")
        self._excel_btn.setObjectName("secondary_btn")
        self._excel_btn.setFixedWidth(120)
        self._excel_btn.clicked.connect(self._browse_excel)
        excel_row.addWidget(self._excel_btn)
        al.addLayout(excel_row)

        sep_row = QtWidgets.QHBoxLayout()
        sep_row.setSpacing(12)
        sep_row.addWidget(make_label("Field separator:", color=TEXT_SEC))
        self._sep_group = QtWidgets.QButtonGroup(self)
        self._sep_pipe = QtWidgets.QRadioButton('  "|"  pipe')
        self._sep_semi = QtWidgets.QRadioButton('  ";"  semicolon')
        self._sep_pipe.setChecked(True)
        self._sep_group.addButton(self._sep_pipe)
        self._sep_group.addButton(self._sep_semi)
        sep_row.addWidget(self._sep_pipe)
        sep_row.addWidget(self._sep_semi)
        sep_row.addStretch()
        al.addLayout(sep_row)

        self._append_widget.hide()
        ops_layout.addWidget(self._append_widget)

        # ── Reformat options ──
        ops_layout.addSpacing(4)
        ops_layout.addWidget(self._radio_reformat)
        self._reformat_widget = QtWidgets.QWidget()
        rfl = QtWidgets.QVBoxLayout(self._reformat_widget)
        rfl.setContentsMargins(20, 4, 0, 4)
        rfl.setSpacing(10)

        self._reformat_mode_group = QtWidgets.QButtonGroup(self)
        self._rf_radio_linearize  = QtWidgets.QRadioButton(
            "Multi-line → Single line  (linearize / unwrap)")
        self._rf_radio_wrap       = QtWidgets.QRadioButton(
            "Single line → Multi-line  (wrap / fold)")
        self._rf_radio_linearize.setChecked(True)
        self._reformat_mode_group.addButton(self._rf_radio_linearize)
        self._reformat_mode_group.addButton(self._rf_radio_wrap)
        self._reformat_mode_group.buttonClicked.connect(self._on_reformat_mode_changed)
        rfl.addWidget(self._rf_radio_linearize)
        rfl.addWidget(self._rf_radio_wrap)

        self._wrap_cols_row = QtWidgets.QWidget()
        wcr = QtWidgets.QHBoxLayout(self._wrap_cols_row)
        wcr.setContentsMargins(0, 0, 0, 0)
        wcr.setSpacing(10)
        wcr.addWidget(make_label("Nucleotides per row:", color=TEXT_SEC))
        self._wrap_cols_spin = QtWidgets.QSpinBox()
        self._wrap_cols_spin.setRange(10, 10000)
        self._wrap_cols_spin.setValue(80)
        self._wrap_cols_spin.setFixedWidth(90)
        wcr.addWidget(self._wrap_cols_spin)
        wcr.addStretch()
        self._wrap_cols_row.hide()
        rfl.addWidget(self._wrap_cols_row)

        self._reformat_widget.hide()
        ops_layout.addWidget(self._reformat_widget)

        # ── Sort options ──
        ops_layout.addSpacing(4)
        ops_layout.addWidget(self._radio_sort)

        # ── Sort options (hidden until selected) ──
        self._sort_widget = QtWidgets.QWidget()
        svl = QtWidgets.QVBoxLayout(self._sort_widget)
        svl.setContentsMargins(20, 4, 0, 4)
        svl.setSpacing(10)

        sep_note = make_label(
            "If headers contain a separator, each resulting field can be used as a sort key.",
            color=TEXT_SEC, size=15,
        )
        sep_note.setWordWrap(True)
        svl.addWidget(sep_note)

        sort_sep_row = QtWidgets.QHBoxLayout()
        sort_sep_row.setSpacing(12)
        sort_sep_row.addWidget(make_label("Field separator:", color=TEXT_SEC))
        self._sort_sep_group = QtWidgets.QButtonGroup(self)
        self._sort_sep_none = QtWidgets.QRadioButton("None (whole header)")
        self._sort_sep_pipe = QtWidgets.QRadioButton('"|"  pipe')
        self._sort_sep_semi = QtWidgets.QRadioButton('";"  semicolon')
        self._sort_sep_custom = QtWidgets.QRadioButton("Custom:")
        self._sort_sep_none.setChecked(True)
        for rb in (self._sort_sep_none, self._sort_sep_pipe, self._sort_sep_semi, self._sort_sep_custom):
            self._sort_sep_group.addButton(rb)
            sort_sep_row.addWidget(rb)
        self._sort_sep_custom_edit = QtWidgets.QLineEdit()
        self._sort_sep_custom_edit.setFixedWidth(64)
        self._sort_sep_custom_edit.setPlaceholderText("e.g. _")
        self._sort_sep_custom_edit.setMaxLength(5)
        self._sort_sep_custom_edit.hide()
        sort_sep_row.addWidget(self._sort_sep_custom_edit)
        sort_sep_row.addStretch()
        svl.addLayout(sort_sep_row)
        self._sort_sep_group.buttonClicked.connect(self._on_sort_sep_changed)
        self._sort_sep_custom_edit.textChanged.connect(self._update_sort_preview)

        # ── Field preview card ──
        self._sort_preview_frame = QtWidgets.QFrame()
        self._sort_preview_frame.setObjectName("sort_preview_card")
        self._sort_preview_frame.setStyleSheet(f"""
            QFrame#sort_preview_card {{
                background: {GRAY_BG};
                border: 1px solid {GRAY_LINE};
                border-radius: 8px;
            }}
        """)
        pfl = QtWidgets.QVBoxLayout(self._sort_preview_frame)
        pfl.setContentsMargins(12, 8, 12, 10)
        pfl.setSpacing(6)
        pfl.addWidget(make_label("Field preview  (first loaded sequence):", size=15, color=TEXT_SEC))
        self._sort_preview_lbl = QtWidgets.QLabel("—  no files loaded")
        self._sort_preview_lbl.setStyleSheet(
            f"font-family:'Courier New',Consolas,monospace; font-size:14px; color:{TEXT_PRI};"
            f" background:transparent;"
        )
        self._sort_preview_lbl.setWordWrap(True)
        pfl.addWidget(self._sort_preview_lbl)
        svl.addWidget(self._sort_preview_frame)

        svl.addWidget(make_label("Sort levels (applied top to bottom):", color=TEXT_SEC))

        self._sort_levels_container = QtWidgets.QWidget()
        self._sort_levels_layout = QtWidgets.QVBoxLayout(self._sort_levels_container)
        self._sort_levels_layout.setContentsMargins(0, 0, 0, 0)
        self._sort_levels_layout.setSpacing(6)
        svl.addWidget(self._sort_levels_container)

        self._sort_level_rows: list = []   # list of (widget, field_spin, order_combo)
        self._add_sort_level()             # start with one level

        add_level_btn = QtWidgets.QPushButton("+ Add sort level")
        add_level_btn.setObjectName("secondary_btn")
        add_level_btn.setFixedWidth(180)
        add_level_btn.clicked.connect(self._add_sort_level)
        svl.addWidget(add_level_btn)

        self._sort_widget.hide()
        ops_layout.addWidget(self._sort_widget)

        self._layout.addWidget(ops_box)

        # ── Status + progress ──
        self._status_lbl = make_label("", color=TEXT_SEC)
        self._layout.addWidget(self._status_lbl)

        self._progress_bar = QtWidgets.QProgressBar()
        self._progress_bar.setRange(0, 0)
        self._progress_bar.setFixedHeight(6)
        self._progress_bar.hide()
        self._layout.addWidget(self._progress_bar)

        self._layout.addStretch()

        # ── Operation log (outside scroll so it sits between scroll and footer) ──
        self._log_edit = QtWidgets.QPlainTextEdit()
        self._log_edit.setReadOnly(True)
        self._log_edit.setFixedHeight(110)
        self._log_edit.setStyleSheet(f"""
            QPlainTextEdit {{
                background: {GRAY_BG};
                border: 1px solid {GRAY_LINE};
                border-top: 1px solid {GRAY_LINE};
                border-radius: 0px;
                font-family: 'Courier New', Consolas, monospace;
                font-size: 16px;
                color: {TEXT_PRI};
                padding: 6px;
            }}
        """)
        self._log_edit.hide()
        outer.addWidget(self._log_edit)

        # ── Footer ──
        footer = QtWidgets.QWidget()
        footer.setObjectName("fasta_tools_footer")
        footer.setStyleSheet(f"""
            QWidget#fasta_tools_footer {{
                background: {GRAY_CARD};
                border-top: 1px solid {GRAY_LINE};
            }}
        """)
        fl = QtWidgets.QHBoxLayout(footer)
        fl.setContentsMargins(20, 10, 20, 10)
        fl.setSpacing(8)

        self._clear_btn = QtWidgets.QPushButton("Clear")
        self._clear_btn.setObjectName("danger_btn")
        self._clear_btn.setFixedHeight(44)
        self._clear_btn.setFixedWidth(140)
        self._clear_btn.clicked.connect(self._clear)
        fl.addWidget(self._clear_btn)

        self._open_folder_btn = QtWidgets.QPushButton("Open folder  📂")
        self._open_folder_btn.setObjectName("secondary_btn")
        self._open_folder_btn.setFixedHeight(44)
        self._open_folder_btn.hide()
        self._open_folder_btn.clicked.connect(self._open_output_folder)
        fl.addWidget(self._open_folder_btn)
        fl.addStretch()

        self._run_btn = QtWidgets.QPushButton("Run  →")
        self._run_btn.setObjectName("primary_btn")
        self._run_btn.setFixedHeight(44)
        self._run_btn.setFixedWidth(300)
        self._run_btn.clicked.connect(self._on_run_clicked)
        self._set_run_enabled(False)
        fl.addWidget(self._run_btn)
        outer.addWidget(footer)

        self._worker: Optional[_FastaToolsWorker] = None
        # Detached-but-maybe-running workers, kept referenced so they aren't GC'd
        # mid-run (which crashes Qt). Purged in _retire_worker.
        self._retired_workers: set = set()
        self._files: list = []
        self._last_outputs: list = []

    # ── helpers ──────────────────────────────────────────────────────────

    @staticmethod
    def _make_section_lbl(text):
        lbl = make_section_label(text)
        lbl.setStyleSheet(
            f"font-size:17px; font-weight:600; color:{BLUE}; letter-spacing:0.4px;"
        )
        return lbl

    @staticmethod
    def _make_clear_btn():
        btn = QtWidgets.QPushButton("✕")
        btn.setFixedSize(26, 26)
        btn.setToolTip("Clear")
        btn.setStyleSheet(
            f"QPushButton {{ background-color: transparent; color: {TEXT_HINT};"
            f" border: none; border-radius: 5px; font-size: 13px; font-weight: bold; }}"
            f"QPushButton:hover {{ background-color: {RED_LT}; color: {RED}; }}"
            f"QPushButton:pressed {{ background-color: {RED}; color: white; }}"
        )
        return btn

    def _set_run_enabled(self, enabled: bool):
        self._run_btn.setEnabled(enabled)
        if enabled:
            self._run_btn.setStyleSheet(
                f"QPushButton {{ background-color:{BLUE}; color:white; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
                f"QPushButton:hover {{ background-color:#0C4A82; }}"
            )
        else:
            self._run_btn.setStyleSheet(
                f"QPushButton {{ background-color:{GRAY_LINE}; color:{TEXT_HINT}; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
            )

    def _retire_worker(self):
        """Detach the current worker: disconnect its slots so a late finish can't
        touch the UI, and keep a reference until its thread exits so it is never
        garbage-collected while still running (which crashes Qt)."""
        self._retired_workers = {w for w in self._retired_workers if w.isRunning()}
        w = self._worker
        self._worker = None
        if w is None:
            return
        for sig in (w.progress, w.log_line, w.finished, w.error):
            try:
                sig.disconnect()
            except (TypeError, RuntimeError):
                pass
        if w.isRunning():
            w.stop()
            self._retired_workers.add(w)

    # ── slots ─────────────────────────────────────────────────────────────

    def _on_files(self, files: list):
        self._files = files
        self._set_run_enabled(bool(files))
        self._status_lbl.setText("")
        self._status_lbl.setStyleSheet("")
        self._update_sort_preview()
        self._update_ff_preview()

    def _on_operation_changed(self, _btn):
        self._grep_widget.setVisible(self._radio_grep.isChecked())
        self._filter_widget.setVisible(self._radio_filter_fields.isChecked())
        self._append_widget.setVisible(self._radio_append.isChecked())
        self._reformat_widget.setVisible(self._radio_reformat.isChecked())
        self._sort_widget.setVisible(self._radio_sort.isChecked())
        if self._radio_sort.isChecked():
            self._update_sort_preview()
        if self._radio_filter_fields.isChecked():
            self._update_ff_preview()
        self._log_edit.clear()
        self._log_edit.hide()
        self._status_lbl.setText("")
        self._status_lbl.setStyleSheet("")
        self._open_folder_btn.hide()

    def _on_reformat_mode_changed(self, _btn):
        self._wrap_cols_row.setVisible(self._rf_radio_wrap.isChecked())

    def _on_sort_sep_changed(self, *_):
        self._sort_sep_custom_edit.setVisible(self._sort_sep_custom.isChecked())
        self._update_sort_preview()

    def _update_sort_preview(self, *_):
        if not self._files:
            self._sort_preview_lbl.setText("—  no files loaded")
            return
        path = self._files[0]
        try:
            import gzip as _gz
            opener = _gz.open if path.lower().endswith(".gz") else open
            header = ""
            with opener(path, "rt", encoding="utf-8", errors="replace") as fh:
                for line in fh:
                    if line.startswith(">"):
                        header = line[1:].rstrip()
                        break
        except Exception:
            self._sort_preview_lbl.setText("—  could not read file")
            return
        if not header:
            self._sort_preview_lbl.setText("—  no sequences found")
            return
        sep = ("|" if self._sort_sep_pipe.isChecked()
               else ";" if self._sort_sep_semi.isChecked()
               else self._sort_sep_custom_edit.text() if self._sort_sep_custom.isChecked()
               else "")
        parts = header.split(sep) if sep and sep in header else [header]
        self._sort_preview_lbl.setText(
            "\n".join(f"Field {i:>2}:  {p.strip()}" for i, p in enumerate(parts, 1))
        )
        self._sort_field_count = len(parts)
        self._update_field_spin_limits()

    def _update_field_spin_limits(self):
        for row in self._sort_level_rows:
            field_spin = row[1]
            field_spin.setMaximum(self._sort_field_count)
            if field_spin.value() > self._sort_field_count:
                field_spin.setValue(self._sort_field_count)

    def _add_sort_level(self):
        n = len(self._sort_level_rows) + 1

        row_w = QtWidgets.QWidget()
        rl = QtWidgets.QHBoxLayout(row_w)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(8)

        lbl = make_label(f"Level {n}:", color=TEXT_SEC)
        lbl.setFixedWidth(60)
        rl.addWidget(lbl)

        rl.addWidget(make_label("Field:", color=TEXT_SEC))
        field_spin = QtWidgets.QSpinBox()
        field_spin.setRange(1, self._sort_field_count)
        field_spin.setValue(min(n, self._sort_field_count))
        field_spin.setFixedWidth(70)
        rl.addWidget(field_spin)

        rl.addSpacing(8)
        rl.addWidget(make_label("Order:", color=TEXT_SEC))
        order_combo = QtWidgets.QComboBox()
        order_combo.addItem("Ascending  ↑", "asc")
        order_combo.addItem("Descending ↓", "desc")
        order_combo.setFixedWidth(160)
        rl.addWidget(order_combo)

        if n > 1:
            remove_btn = QtWidgets.QPushButton("✕")
            remove_btn.setFixedSize(28, 28)
            remove_btn.setStyleSheet(
                "QPushButton { background:transparent; border:1px solid #CCC;"
                " border-radius:4px; color:#888; font-size:11px; }"
                "QPushButton:hover { background:#FEE2E2; border-color:#EF4444; color:#EF4444; }"
            )
            remove_btn.clicked.connect(lambda _, w=row_w: self._remove_sort_level(w))
            rl.addWidget(remove_btn)

        rl.addStretch()
        self._sort_levels_layout.addWidget(row_w)
        self._sort_level_rows.append((row_w, field_spin, order_combo))

    def _remove_sort_level(self, row_widget):
        for i, (w, _fs, _oc) in enumerate(self._sort_level_rows):
            if w is row_widget:
                self._sort_level_rows.pop(i)
                self._sort_levels_layout.removeWidget(w)
                w.deleteLater()
                self._renumber_sort_levels()
                break

    def _renumber_sort_levels(self):
        for i, (w, _fs, _oc) in enumerate(self._sort_level_rows):
            lbl = w.layout().itemAt(0).widget()
            if isinstance(lbl, QtWidgets.QLabel):
                lbl.setText(f"Level {i + 1}:")

    def _on_ff_sep_changed(self, *_):
        self._ff_sep_custom_edit.setVisible(self._ff_sep_custom.isChecked())
        self._update_ff_preview()

    def _update_ff_preview(self, *_):
        if not self._files:
            self._ff_preview_lbl.setText("—  no files loaded")
            return
        path = self._files[0]
        try:
            import gzip as _gz
            opener = _gz.open if path.lower().endswith(".gz") else open
            header = ""
            with opener(path, "rt", encoding="utf-8", errors="replace") as fh:
                for line in fh:
                    if line.startswith(">"):
                        header = line[1:].rstrip()
                        break
        except Exception:
            self._ff_preview_lbl.setText("—  could not read file")
            return
        if not header:
            self._ff_preview_lbl.setText("—  no sequences found")
            return
        sep = ("|" if self._ff_sep_pipe.isChecked()
               else ";" if self._ff_sep_semi.isChecked()
               else self._ff_sep_custom_edit.text() if self._ff_sep_custom.isChecked()
               else "|")
        parts = header.split(sep) if sep and sep in header else [header]
        self._ff_preview_lbl.setText(
            "\n".join(f"Field {i:>2}:  {p.strip()}" for i, p in enumerate(parts, 1))
        )
        self._ff_field_count = len(parts)
        self._update_ff_spin_limits()

    def _update_ff_spin_limits(self):
        for _w, fs, _oc, _ve in self._filter_criterion_rows:
            fs.setMaximum(self._ff_field_count)
            if fs.value() > self._ff_field_count:
                fs.setValue(self._ff_field_count)

    def _add_filter_criterion(self):
        _REMOVE_STYLE = (
            "QPushButton { background:transparent; border:1px solid #CCC;"
            " border-radius:4px; color:#888; font-size:11px; }"
            "QPushButton:hover { background:#FEE2E2; border-color:#EF4444; color:#EF4444; }"
        )
        n = len(self._filter_criterion_rows) + 1
        row_w = QtWidgets.QWidget()
        rl = QtWidgets.QHBoxLayout(row_w)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(6)

        rl.addWidget(make_label("Field:", color=TEXT_SEC))
        field_spin = QtWidgets.QSpinBox()
        field_spin.setRange(1, self._ff_field_count)
        field_spin.setValue(min(n, self._ff_field_count))
        field_spin.setFixedWidth(70)
        rl.addWidget(field_spin)

        op_combo = QtWidgets.QComboBox()
        for op in ("=", "!=", ">=", "<=", ">", "<", "contains", "not contains"):
            op_combo.addItem(op)
        op_combo.setFixedWidth(130)
        rl.addWidget(op_combo)

        value_edit = QtWidgets.QLineEdit()
        value_edit.setPlaceholderText("value")
        value_edit.setFixedWidth(140)
        rl.addWidget(value_edit)

        if self._filter_criterion_rows:
            remove_btn = QtWidgets.QPushButton("✕")
            remove_btn.setFixedSize(28, 28)
            remove_btn.setStyleSheet(_REMOVE_STYLE)
            remove_btn.clicked.connect(lambda _, w=row_w: self._remove_filter_criterion(w))
            rl.addWidget(remove_btn)

        rl.addStretch()
        self._filter_criteria_layout.addWidget(row_w)
        self._filter_criterion_rows.append((row_w, field_spin, op_combo, value_edit))

    def _remove_filter_criterion(self, row_widget):
        for i, (w, _fs, _oc, _ve) in enumerate(self._filter_criterion_rows):
            if w is row_widget:
                self._filter_criterion_rows.pop(i)
                self._filter_criteria_layout.removeWidget(w)
                w.deleteLater()
                break

    def _on_grep_source_changed(self, state):
        use_file = (state == QtCore.Qt.Checked)
        self._grep_pattern_row.setVisible(not use_file)
        self._grep_file_row.setVisible(use_file)
        self._grep_file_note.setVisible(use_file)

    def _browse_grep_file(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select patterns file", "", "Text files (*.txt);;All files (*)"
        )
        if path:
            self._grep_file_edit.setText(path)

    def _browse_excel(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select Excel file", "", "Excel files (*.xlsx *.xls)"
        )
        if path:
            self._excel_edit.setText(path)

    def _clear(self):
        self._retire_worker()

        # ── Files / status ──
        self._drop.clear()
        self._files = []
        self._last_outputs = []
        self._status_lbl.setText("")
        self._status_lbl.setStyleSheet("")
        self._progress_bar.hide()
        self._log_edit.clear()
        self._log_edit.hide()
        self._open_folder_btn.hide()
        self._set_run_enabled(False)
        self._merge_chk.setChecked(False)

        # ── Operation: back to "unique" ──
        self._radio_unique.setChecked(True)
        self._grep_widget.hide()
        self._filter_widget.hide()
        self._append_widget.hide()
        self._reformat_widget.hide()
        self._sort_widget.hide()

        # ── Grep options ──
        self._grep_use_file_chk.setChecked(False)
        self._grep_pattern_edit.clear()
        self._grep_file_edit.clear()
        self._grep_regex_chk.setChecked(False)
        self._grep_pattern_row.show()
        self._grep_file_row.hide()
        self._grep_file_note.hide()

        # ── Append options ──
        self._excel_edit.clear()
        self._sep_pipe.setChecked(True)

        # ── Reformat options ──
        self._rf_radio_linearize.setChecked(True)
        self._wrap_cols_spin.setValue(80)
        self._wrap_cols_row.hide()

        # ── Sort options ──
        self._sort_sep_none.setChecked(True)
        self._sort_sep_custom_edit.clear()
        self._sort_sep_custom_edit.hide()
        # Remove all levels beyond the first
        while len(self._sort_level_rows) > 1:
            w, _fs, _oc = self._sort_level_rows[-1]
            self._sort_level_rows.pop()
            self._sort_levels_layout.removeWidget(w)
            w.deleteLater()
        # Reset first level
        if self._sort_level_rows:
            _w, fs, oc = self._sort_level_rows[0]
            fs.setValue(1)
            oc.setCurrentIndex(0)
        self._update_sort_preview()

        # ── Filter by fields options ──
        self._ff_sep_pipe.setChecked(True)
        self._ff_sep_custom_edit.clear()
        self._ff_sep_custom_edit.hide()
        while self._filter_criterion_rows:
            w, _fs, _oc, _ve = self._filter_criterion_rows.pop()
            self._filter_criteria_layout.removeWidget(w)
            w.deleteLater()
        self._add_filter_criterion()

    def _on_run_clicked(self):
        if not self._files:
            return

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"ont-barcoder_{ts}_fastools"
        auto_dir = os.path.join(_get_base_dir(), "output", folder_name)

        out_dir = self._ask_output_dir(auto_dir, folder_name)
        if out_dir is None:
            return

        operation = (
            "unique"         if self._radio_unique.isChecked()         else
            "identical"      if self._radio_identical.isChecked()      else
            "grep"           if self._radio_grep.isChecked()           else
            "filter_fields"  if self._radio_filter_fields.isChecked()  else
            "append"         if self._radio_append.isChecked()         else
            "reformat"       if self._radio_reformat.isChecked()       else
            "sort"
        )

        params: dict = {"merge": self._merge_chk.isChecked()}
        if operation == "grep":
            params["use_regex"] = self._grep_regex_chk.isChecked()
            if self._grep_use_file_chk.isChecked():
                params["pattern_file"] = self._grep_file_edit.text().strip()
            else:
                params["pattern"] = self._grep_pattern_edit.text().strip()
        elif operation == "filter_fields":
            if self._ff_sep_semi.isChecked():
                params["separator"] = ";"
            elif self._ff_sep_custom.isChecked():
                custom_sep = self._ff_sep_custom_edit.text()
                if not custom_sep:
                    QtWidgets.QMessageBox.warning(
                        self, "Filter configuration",
                        "Custom separator is empty.\n"
                        "Please enter a separator character or choose a different option."
                    )
                    return
                params["separator"] = custom_sep
            else:
                params["separator"] = "|"
            criteria = [
                {"field": fs.value(), "op": oc.currentText(), "value": ve.text().strip()}
                for _w, fs, oc, ve in self._filter_criterion_rows
                if ve.text().strip()
            ]
            if not criteria:
                QtWidgets.QMessageBox.warning(
                    self, "Filter configuration",
                    "No valid criteria.\n"
                    "Please fill in at least one value."
                )
                return
            # Numeric operators need a numeric value; otherwise the criterion
            # would fall into the text branch (which has no case for them) and
            # silently pass every sequence.
            for c in criteria:
                if c["op"] in (">=", "<=", ">", "<"):
                    try:
                        float(c["value"])
                    except ValueError:
                        QtWidgets.QMessageBox.warning(
                            self, "Filter configuration",
                            f"Operator '{c['op']}' (field {c['field']}) requires a "
                            f"numeric value, but got “{c['value']}”.\n\n"
                            "Enter a number or use a text operator "
                            "(=, !=, contains, not contains)."
                        )
                        return
            params["criteria"] = criteria
        elif operation == "append":
            params["excel_file"] = self._excel_edit.text().strip()
            params["separator"]  = "|" if self._sep_pipe.isChecked() else ";"
        elif operation == "reformat":
            params["mode"]      = "wrap" if self._rf_radio_wrap.isChecked() else "linearize"
            params["wrap_cols"] = self._wrap_cols_spin.value()
        elif operation == "sort":
            if self._sort_sep_pipe.isChecked():
                params["separator"] = "|"
            elif self._sort_sep_semi.isChecked():
                params["separator"] = ";"
            elif self._sort_sep_custom.isChecked():
                custom_sep = self._sort_sep_custom_edit.text()
                if not custom_sep:
                    QtWidgets.QMessageBox.warning(
                        self, "Sort configuration",
                        "Custom separator is empty.\n"
                        "Please enter a separator character or choose a different option."
                    )
                    return
                params["separator"] = custom_sep
            else:
                params["separator"] = ""
            params["levels"] = [
                {"field": fs.value(), "order": oc.currentData()}
                for _, fs, oc in self._sort_level_rows
            ]
            # Validate: no separator means only 1 field exists
            if not params["separator"]:
                bad = [lv for lv in params["levels"] if lv["field"] > 1]
                if bad:
                    QtWidgets.QMessageBox.warning(
                        self, "Sort configuration",
                        "Separator is set to None (whole header = 1 field).\n"
                        "All sort levels must use Field 1.\n\n"
                        "Please select a separator or set all levels to Field 1."
                    )
                    return

        self._progress_bar.show()
        self._status_lbl.setStyleSheet("")
        self._status_lbl.setText("Running…")
        self._set_run_enabled(False)
        self._log_edit.clear()
        self._log_edit.show()

        self._retire_worker()
        self._worker = _FastaToolsWorker(self._files, operation, params, out_dir)
        self._worker.progress.connect(self._on_progress)
        self._worker.log_line.connect(self._on_log_line)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._worker.start()

    def _ask_output_dir(self, auto_dir: str, folder_name: str) -> Optional[str]:
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Output folder")
        dlg.setMinimumWidth(480)
        dlg.setStyleSheet(f"""
            QDialog {{ background-color: {GRAY_CARD}; }}
            QLabel {{ color: {TEXT_PRI}; background-color: transparent; }}
            QRadioButton {{
                color: {TEXT_PRI}; background-color: transparent;
                font-size: 15px; padding: 6px 0;
            }}
            QRadioButton::indicator {{ width: 16px; height: 16px; }}
            QPushButton {{
                border-radius: 8px; padding: 8px 20px;
                font-size: 15px; font-weight: 500;
            }}
            #dlg_ok_btn {{ background-color: {BLUE}; color: white; border: none; }}
            #dlg_ok_btn:hover {{ background-color: #0C4A82; }}
            #dlg_cancel_btn {{
                background-color: transparent; color: {BLUE};
                border: 1px solid {BLUE};
            }}
            #dlg_cancel_btn:hover {{ background-color: {BLUE_LIGHT}; }}
        """)

        vlay = QtWidgets.QVBoxLayout(dlg)
        vlay.setSpacing(16)
        vlay.setContentsMargins(24, 24, 24, 20)

        title_lbl = QtWidgets.QLabel("Where to save the results?")
        title_lbl.setStyleSheet(f"font-size:17px; font-weight:700; color:{TEXT_PRI};")
        vlay.addWidget(title_lbl)

        radio_auto = QtWidgets.QRadioButton(
            f"Automatic folder (recommended)\n  …/output/{folder_name}/"
        )
        radio_auto.setChecked(True)
        radio_custom = QtWidgets.QRadioButton("Select folder manually")
        vlay.addWidget(radio_auto)
        vlay.addWidget(radio_custom)
        vlay.addSpacing(8)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_cancel.setObjectName("dlg_cancel_btn")
        btn_cancel.setFixedHeight(38)
        btn_ok = QtWidgets.QPushButton("Run")
        btn_ok.setObjectName("dlg_ok_btn")
        btn_ok.setFixedHeight(38)
        btn_ok.setDefault(True)
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addSpacing(8)
        btn_row.addWidget(btn_ok)
        vlay.addLayout(btn_row)

        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return None

        if radio_auto.isChecked():
            return auto_dir
        parent_dir = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select output folder"
        )
        if not parent_dir:
            return None
        return os.path.join(parent_dir, folder_name)

    def _on_progress(self, msg: str):
        self._status_lbl.setText(msg)

    def _on_log_line(self, line: str):
        self._log_edit.appendPlainText(line)

    def _on_finished(self, outputs: list):
        self._progress_bar.hide()
        self._set_run_enabled(True)
        self._last_outputs = outputs
        self._status_lbl.setStyleSheet(f"color:{GREEN};")
        self._status_lbl.setText(f"Done — {len(outputs)} file(s) saved.")
        self._log_edit.appendPlainText(
            f"─── {len(outputs)} file(s) saved to: "
            f"{os.path.dirname(outputs[0]) if outputs else '—'}"
        )
        if outputs:
            self._open_folder_btn.show()

    def _on_error(self, msg: str):
        self._progress_bar.hide()
        self._set_run_enabled(True)
        self._status_lbl.setStyleSheet(f"color:{RED};")
        self._status_lbl.setText(f"Error: {msg}")
        self._log_edit.appendPlainText(f"ERROR: {msg}")

    def _open_output_folder(self):
        if not self._last_outputs:
            return
        folder = os.path.dirname(self._last_outputs[0])
        if os.path.isdir(folder):
            QtGui.QDesktopServices.openUrl(
                QtCore.QUrl.fromLocalFile(folder)
            )


# ═══════════════════════════════════════════════════════════════════════════
# BOLD → BLAST FORMAT
# ═══════════════════════════════════════════════════════════════════════════


class _BoldFormatWorker(QtCore.QThread):
    """Apply the BLAST-results visual format to BOLD Barcode-ID xlsx exports.

    Takes a BOLD "Barcode ID" workbook (a title row, a header row, then hit
    rows grouped by Query ID and sorted by ID% descending) and produces a
    styled copy that mirrors the look of the BLAST results workbook.
    """

    progress = QtCore.pyqtSignal(str)   # short status shown in the label
    log_line = QtCore.pyqtSignal(str)   # detailed per-file result appended to the log
    finished = QtCore.pyqtSignal(list)
    error    = QtCore.pyqtSignal(str)

    # --- Palette taken from the BLAST results workbook -----------------------
    GRID_COLOR  = "FFCCCCCC"      # thin cell borders
    SEP_COLOR   = "FF9AA0A6"      # medium separator between Query ID groups
    HEADER_TEXT = "FFFFFFFF"
    FONT_SIZE   = 10.0

    QUERY_COL_NAME   = "Query ID"
    IDPCT_COL_NAME   = "ID%"
    HITRANK_COL_NAME = "Hit rank"

    # Semantic column groups by header name -> ("header_rgb", "band_rgb").
    _BLUE   = ("FF1A365D", "FFE8F1FB")
    _TEAL   = ("FF0D5E6E", "FFE8F5F6")
    _ORANGE = ("FF7C3200", "FFFEF3E8")

    COLOR_BY_HEADER = {
        "Hit rank":  _BLUE,
        "Query ID":  _BLUE,
        "PID [BIN]": _TEAL,
        "Phylum":    _ORANGE,
        "Class":     _ORANGE,
        "Order":     _ORANGE,
        "Family":    _ORANGE,
        "Subfamily": _ORANGE,
        "Tribe":     _ORANGE,
        "Genus":     _ORANGE,
        "Species":   _ORANGE,
        "Indels":    _TEAL,
        "ID%":       _TEAL,
    }

    WIDTH_BY_HEADER = {
        "Hit rank":  10,
        "Query ID":  55,
        "PID [BIN]": 16,
        "Phylum":    15,
        "Class":     15,
        "Order":     15,
        "Family":    16,
        "Subfamily": 14,
        "Tribe":     12,
        "Genus":     18,
        "Species":   28,
        "Indels":     9,
        "ID%":        9,
    }

    def __init__(self, files: list, max_hits: Optional[int], out_dir: str, parent=None):
        super().__init__(parent)
        self._files    = files
        self._max_hits = max_hits   # None = keep all hits
        self._out_dir  = out_dir
        self._stop     = False

    def stop(self):
        self._stop = True

    # ── core logic (ported from boldID_2_blast_format.py) ──────────────────

    @classmethod
    def _read_bold(cls, path):
        """Return (headers, groups, q_idx, id_idx).

        The input has a title in row 1, headers in row 2 and data from row 3 on.
        Rows keep their original relative order (already sorted by ID% desc).
        """
        from openpyxl import load_workbook

        wb = load_workbook(path)
        ws = wb.active
        max_col = ws.max_column

        headers = [ws.cell(row=2, column=c).value for c in range(1, max_col + 1)]
        if cls.QUERY_COL_NAME not in headers:
            raise ValueError(f"no '{cls.QUERY_COL_NAME}' column found")
        if cls.IDPCT_COL_NAME not in headers:
            raise ValueError(f"no '{cls.IDPCT_COL_NAME}' column found")

        q_idx = headers.index(cls.QUERY_COL_NAME)
        id_idx = headers.index(cls.IDPCT_COL_NAME)

        groups = OrderedDict()
        for r in range(3, ws.max_row + 1):
            row = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
            if row[q_idx] is None and all(v is None for v in row):
                continue  # skip fully blank rows
            groups.setdefault(row[q_idx], []).append(row)

        return headers, groups, q_idx, id_idx

    @classmethod
    def _build(cls, headers, groups, q_idx, id_idx, max_hits):
        """Sort/trim each group and return (out_headers, records, reordered).

        records is a list of (row_values, hit_rank, group_index, is_first_in_group).
        A "Hit rank" column is prepended.
        """
        out_headers = [cls.HITRANK_COL_NAME] + headers
        records = []
        reordered = []

        for gi, (query, rows) in enumerate(groups.items(), start=1):
            def id_key(row):
                v = row[id_idx]
                return v if isinstance(v, (int, float)) else float("-inf")

            vals = [id_key(r) for r in rows]
            if any(a < b for a, b in zip(vals, vals[1:])):
                reordered.append(query)
            rows_sorted = sorted(rows, key=id_key, reverse=True)

            if max_hits is not None:
                rows_sorted = rows_sorted[:max_hits]

            for rank, row in enumerate(rows_sorted, start=1):
                records.append(([rank] + row, rank, gi, rank == 1))

        return out_headers, records, reordered

    @classmethod
    def _write(cls, out_headers, records, out_path):
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter

        def _side(style, rgb):
            return Side(style=style, color=rgb)

        wb = Workbook()
        ws = wb.active
        ws.title = "Barcode ID"

        thin = _side("thin", cls.GRID_COLOR)
        border_thin = Border(top=thin, bottom=thin, left=thin, right=thin)
        border_sep = Border(top=_side("medium", cls.SEP_COLOR),
                            bottom=thin, left=thin, right=thin)

        header_fills = {}
        band_fills = {}
        for name in set(out_headers):
            hrgb, brgb = cls.COLOR_BY_HEADER.get(name, cls._TEAL)
            header_fills[name] = PatternFill("solid", fgColor=hrgb)
            band_fills[name] = PatternFill("solid", fgColor=brgb)

        center = Alignment(horizontal="center", vertical="center")

        # Header row.
        for c, name in enumerate(out_headers, start=1):
            cell = ws.cell(row=1, column=c, value=name)
            cell.fill = header_fills[name]
            cell.font = Font(bold=True, color=cls.HEADER_TEXT, size=cls.FONT_SIZE)
            cell.border = border_thin
            cell.alignment = center

        idpct_col = out_headers.index(cls.IDPCT_COL_NAME) + 1

        # Data rows.
        for i, (values, rank, gi, is_first) in enumerate(records):
            r = i + 2
            banded = (gi % 2 == 1)  # odd groups colored, even groups white
            for c, name in enumerate(out_headers, start=1):
                cell = ws.cell(row=r, column=c, value=values[c - 1])
                cell.font = Font(bold=is_first, size=cls.FONT_SIZE)
                if banded:
                    cell.fill = band_fills[name]
                cell.border = border_sep if (is_first and gi > 1) else border_thin
                if name == cls.HITRANK_COL_NAME:
                    cell.alignment = center
            ws.cell(row=r, column=idpct_col).number_format = "0.00"

        # Column widths, freeze and filter.
        for c, name in enumerate(out_headers, start=1):
            ws.column_dimensions[get_column_letter(c)].width = cls.WIDTH_BY_HEADER.get(name, 15)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(out_headers))}{len(records) + 1}"

        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        wb.save(out_path)

    # ── thread entry point ─────────────────────────────────────────────────

    def run(self):
        try:
            outputs = []
            total = len(self._files)
            for i, path in enumerate(self._files, 1):
                if self._stop:
                    return
                name = os.path.basename(path)
                self.progress.emit(f"Formatting {name}…  ({i}/{total})")
                try:
                    headers, groups, q_idx, id_idx = self._read_bold(path)
                    out_headers, records, reordered = self._build(
                        headers, groups, q_idx, id_idx, self._max_hits)
                    out_path = os.path.join(
                        self._out_dir,
                        f"{os.path.splitext(name)[0]}_formatted.xlsx")
                    self._write(out_headers, records, out_path)
                    outputs.append(out_path)

                    kept = "all" if self._max_hits is None else self._max_hits
                    self.log_line.emit(
                        f"{name}: {len(groups)} Query ID(s) · "
                        f"{len(records)} row(s) (hits/group: {kept}) "
                        f"→ {os.path.basename(out_path)}")
                    if reordered:
                        self.log_line.emit(
                            f"  ⚠ {len(reordered)} group(s) were not sorted by "
                            f"ID% desc and got re-sorted")
                except Exception as exc:
                    self.log_line.emit(f"{name}: ERROR — {exc}")

            if self._stop:
                return
            if not outputs:
                self.error.emit(
                    "No files were formatted. Make sure the input is a BOLD "
                    "Barcode-ID export (title row, header row, then hits).")
                return
            self.finished.emit(outputs)
        except Exception as exc:  # pragma: no cover - defensive
            self.error.emit(str(exc))


class BoldFormatPanel(QtWidgets.QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)

        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── Scroll area ──
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        inner = QtWidgets.QWidget()
        self._layout = QtWidgets.QVBoxLayout(inner)
        self._layout.setContentsMargins(20, 20, 20, 20)
        self._layout.setSpacing(16)
        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        # ── Title ──
        self._layout.addWidget(make_label("BOLD Formatter", size=19, bold=True))
        desc = make_label(
            "Take a BOLD “Barcode ID” .xlsx export and restyle it to match the "
            "BLAST results workbook: a per-group “Hit rank” column, top-N hits "
            "per Query ID, group colors, banded fills, frozen header and filters.",
            color=TEXT_SEC,
        )
        desc.setWordWrap(True)
        self._layout.addWidget(desc)

        # ── Input file (drag & drop + browse) ──
        self._layout.addWidget(self._make_section_lbl("Input file"))

        file_row = QtWidgets.QHBoxLayout()
        file_row.setSpacing(8)
        lbl_in = make_label("BOLD .xlsx:", color=TEXT_SEC)
        lbl_in.setToolTip(
            "BOLD 'Barcode ID' workbook: a title row, a header row with a "
            "'Query ID' and 'ID%' column, then hit rows grouped by Query ID."
        )
        file_row.addWidget(lbl_in)
        self._file_edit = _DragDropLineEdit(accepted_extensions=[".xlsx"])
        self._file_edit.setReadOnly(True)
        self._file_edit.setPlaceholderText("No file selected… (or drag & drop)")
        self._file_edit.textChanged.connect(self._on_file_changed)
        file_row.addWidget(self._file_edit, 1)
        self._browse_btn = QtWidgets.QPushButton("Browse…")
        self._browse_btn.setObjectName("secondary_btn")
        self._browse_btn.setFixedWidth(120)
        self._browse_btn.clicked.connect(self._browse_file)
        file_row.addWidget(self._browse_btn)
        self._layout.addLayout(file_row)

        # ── Options ──
        self._layout.addWidget(self._make_section_lbl("Options"))

        hits_row = QtWidgets.QHBoxLayout()
        hits_row.setSpacing(12)
        lbl_hits = make_label("Hits to keep per Query ID:", color=TEXT_SEC)
        hits_row.addWidget(lbl_hits)
        self._hits_spin = QtWidgets.QSpinBox()
        self._hits_spin.setRange(1, 999)
        self._hits_spin.setValue(5)
        self._hits_spin.setFixedWidth(90)
        hits_row.addWidget(self._hits_spin)
        hits_row.addStretch()
        self._layout.addLayout(hits_row)

        self._keep_all_chk = QtWidgets.QCheckBox("Keep all hits (ignore the limit above)")
        self._keep_all_chk.toggled.connect(
            lambda on: (self._hits_spin.setDisabled(on), lbl_hits.setDisabled(on)))
        self._layout.addWidget(self._keep_all_chk)

        # ── Status + progress ──
        self._status_lbl = make_label("", color=TEXT_SEC)
        self._layout.addWidget(self._status_lbl)

        self._progress_bar = QtWidgets.QProgressBar()
        self._progress_bar.setRange(0, 0)
        self._progress_bar.setFixedHeight(6)
        self._progress_bar.hide()
        self._layout.addWidget(self._progress_bar)

        self._layout.addStretch()

        # ── Operation log ──
        self._log_edit = QtWidgets.QPlainTextEdit()
        self._log_edit.setReadOnly(True)
        self._log_edit.setFixedHeight(110)
        self._log_edit.setStyleSheet(f"""
            QPlainTextEdit {{
                background: {GRAY_BG};
                border: 1px solid {GRAY_LINE};
                border-top: 1px solid {GRAY_LINE};
                border-radius: 0px;
                font-family: 'Courier New', Consolas, monospace;
                font-size: 16px;
                color: {TEXT_PRI};
                padding: 6px;
            }}
        """)
        self._log_edit.hide()
        outer.addWidget(self._log_edit)

        # ── Footer ──
        footer = QtWidgets.QWidget()
        footer.setObjectName("bold_format_footer")
        footer.setStyleSheet(f"""
            QWidget#bold_format_footer {{
                background: {GRAY_CARD};
                border-top: 1px solid {GRAY_LINE};
            }}
        """)
        fl = QtWidgets.QHBoxLayout(footer)
        fl.setContentsMargins(20, 10, 20, 10)
        fl.setSpacing(8)

        self._clear_btn = QtWidgets.QPushButton("Clear")
        self._clear_btn.setObjectName("danger_btn")
        self._clear_btn.setFixedHeight(44)
        self._clear_btn.setFixedWidth(140)
        self._clear_btn.clicked.connect(self._clear)
        fl.addWidget(self._clear_btn)

        self._open_folder_btn = QtWidgets.QPushButton("Open folder  📂")
        self._open_folder_btn.setObjectName("secondary_btn")
        self._open_folder_btn.setFixedHeight(44)
        self._open_folder_btn.hide()
        self._open_folder_btn.clicked.connect(self._open_output_folder)
        fl.addWidget(self._open_folder_btn)
        fl.addStretch()

        self._run_btn = QtWidgets.QPushButton("Format  →")
        self._run_btn.setObjectName("primary_btn")
        self._run_btn.setFixedHeight(44)
        self._run_btn.setFixedWidth(300)
        self._run_btn.clicked.connect(self._on_run_clicked)
        fl.addWidget(self._run_btn)
        self._set_run_enabled(False)
        outer.addWidget(footer)

        self._worker: Optional[_BoldFormatWorker] = None
        self._retired_workers: set = set()
        self._last_outputs: list = []

    # ── helpers ─────────────────────────────────────────────────────────────

    @staticmethod
    def _make_section_lbl(text):
        lbl = make_section_label(text)
        lbl.setStyleSheet(
            f"font-size:17px; font-weight:600; color:{BLUE}; letter-spacing:0.4px;"
        )
        return lbl

    def _set_run_enabled(self, enabled: bool):
        self._run_btn.setEnabled(enabled)
        if enabled:
            self._run_btn.setStyleSheet(
                f"QPushButton {{ background-color:{BLUE}; color:white; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
                f"QPushButton:hover {{ background-color:#0C4A82; }}"
            )
        else:
            self._run_btn.setStyleSheet(
                f"QPushButton {{ background-color:{GRAY_LINE}; color:{TEXT_HINT}; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
            )

    def _retire_worker(self):
        """Detach the current worker so a late finish can't touch the UI, and keep
        a reference until its thread exits so it is never garbage-collected while
        still running (which crashes Qt)."""
        self._retired_workers = {w for w in self._retired_workers if w.isRunning()}
        w = self._worker
        self._worker = None
        if w is None:
            return
        for sig in (w.progress, w.log_line, w.finished, w.error):
            try:
                sig.disconnect()
            except (TypeError, RuntimeError):
                pass
        if w.isRunning():
            w.stop()
            self._retired_workers.add(w)

    # ── slots ────────────────────────────────────────────────────────────────

    def _browse_file(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select BOLD Barcode-ID .xlsx", "", "Excel files (*.xlsx)"
        )
        if path:
            self._file_edit.setText(path)

    def _on_file_changed(self, text: str):
        self._set_run_enabled(bool(text.strip()))
        self._status_lbl.setText("")
        self._status_lbl.setStyleSheet("")
        self._open_folder_btn.hide()

    def _clear(self):
        self._file_edit.clear()
        self._hits_spin.setValue(5)
        self._keep_all_chk.setChecked(False)
        self._status_lbl.setText("")
        self._status_lbl.setStyleSheet("")
        self._progress_bar.hide()
        self._log_edit.clear()
        self._log_edit.hide()
        self._open_folder_btn.hide()
        self._set_run_enabled(False)

    def _on_run_clicked(self):
        path = self._file_edit.text().strip()
        if not path:
            return
        if not os.path.isfile(path):
            QtWidgets.QMessageBox.warning(
                self, "Input file", f"File not found:\n{path}")
            return

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"ont-barcoder_{ts}_bold_format"
        auto_dir = os.path.join(_get_base_dir(), "output", folder_name)

        out_dir = self._ask_output_dir(auto_dir, folder_name)
        if out_dir is None:
            return

        max_hits = None if self._keep_all_chk.isChecked() else self._hits_spin.value()

        self._progress_bar.show()
        self._status_lbl.setStyleSheet("")
        self._status_lbl.setText("Formatting…")
        self._set_run_enabled(False)
        self._log_edit.clear()
        self._log_edit.show()
        self._open_folder_btn.hide()

        self._retire_worker()
        self._worker = _BoldFormatWorker([path], max_hits, out_dir)
        self._worker.progress.connect(self._on_progress)
        self._worker.log_line.connect(self._on_log_line)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._worker.start()

    def _ask_output_dir(self, auto_dir: str, folder_name: str) -> Optional[str]:
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Output folder")
        dlg.setMinimumWidth(480)
        dlg.setStyleSheet(f"""
            QDialog {{ background-color: {GRAY_CARD}; }}
            QLabel {{ color: {TEXT_PRI}; background-color: transparent; }}
            QRadioButton {{
                color: {TEXT_PRI}; background-color: transparent;
                font-size: 15px; padding: 6px 0;
            }}
            QRadioButton::indicator {{ width: 16px; height: 16px; }}
            QPushButton {{
                border-radius: 8px; padding: 8px 20px;
                font-size: 15px; font-weight: 500;
            }}
            #dlg_ok_btn {{ background-color: {BLUE}; color: white; border: none; }}
            #dlg_ok_btn:hover {{ background-color: #0C4A82; }}
            #dlg_cancel_btn {{
                background-color: transparent; color: {BLUE};
                border: 1px solid {BLUE};
            }}
            #dlg_cancel_btn:hover {{ background-color: {BLUE_LIGHT}; }}
        """)

        vlay = QtWidgets.QVBoxLayout(dlg)
        vlay.setSpacing(16)
        vlay.setContentsMargins(24, 24, 24, 20)

        title_lbl = QtWidgets.QLabel("Where to save the results?")
        title_lbl.setStyleSheet(f"font-size:17px; font-weight:700; color:{TEXT_PRI};")
        vlay.addWidget(title_lbl)

        radio_auto = QtWidgets.QRadioButton(
            f"Automatic folder (recommended)\n  …/output/{folder_name}/"
        )
        radio_auto.setChecked(True)
        radio_custom = QtWidgets.QRadioButton("Select folder manually")
        vlay.addWidget(radio_auto)
        vlay.addWidget(radio_custom)
        vlay.addSpacing(8)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_cancel.setObjectName("dlg_cancel_btn")
        btn_cancel.setFixedHeight(38)
        btn_ok = QtWidgets.QPushButton("Run")
        btn_ok.setObjectName("dlg_ok_btn")
        btn_ok.setFixedHeight(38)
        btn_ok.setDefault(True)
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addSpacing(8)
        btn_row.addWidget(btn_ok)
        vlay.addLayout(btn_row)

        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return None

        if radio_auto.isChecked():
            return auto_dir
        parent_dir = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select output folder"
        )
        if not parent_dir:
            return None
        return os.path.join(parent_dir, folder_name)

    def _on_progress(self, msg: str):
        self._status_lbl.setText(msg)

    def _on_log_line(self, line: str):
        self._log_edit.appendPlainText(line)

    def _on_finished(self, outputs: list):
        self._progress_bar.hide()
        self._set_run_enabled(True)
        self._last_outputs = outputs
        self._status_lbl.setStyleSheet(f"color:{GREEN};")
        self._status_lbl.setText(f"Done — {len(outputs)} file(s) saved.")
        self._log_edit.appendPlainText(
            f"─── {len(outputs)} file(s) saved to: "
            f"{os.path.dirname(outputs[0]) if outputs else '—'}"
        )
        if outputs:
            self._open_folder_btn.show()

    def _on_error(self, msg: str):
        self._progress_bar.hide()
        self._set_run_enabled(True)
        self._status_lbl.setStyleSheet(f"color:{RED};")
        self._status_lbl.setText(f"Error: {msg}")
        self._log_edit.appendPlainText(f"ERROR: {msg}")

    def _open_output_folder(self):
        if not self._last_outputs:
            return
        folder = os.path.dirname(self._last_outputs[0])
        if os.path.isdir(folder):
            QtGui.QDesktopServices.openUrl(
                QtCore.QUrl.fromLocalFile(folder)
            )


# ═══════════════════════════════════════════════════════════════════════════
# GENBANK BATCH SEARCH
# ═══════════════════════════════════════════════════════════════════════════


# Columns offered for the customizable summary report (same idea as the BOLD
# app's batch-search "report" block): per term, how many distinct values were
# found for a column, and/or which ones. Keyed by the field names produced by
# _GenbankBatchWorker._iter_records(); shared between the worker (headers) and
# the panel (checkbox grid) so both stay in sync.
_GENBANK_REPORT_COLUMNS = [
    ("marker",   "Marker/Gene"),
    ("moltype",  "Moltype"),
    ("division", "Division"),
    ("country",  "Country"),
]


class _GenbankBatchWorker(QtCore.QThread):
    """Batch-search NCBI GenBank (nucleotide) with up to two crossed criteria.

    Python replacement for the old GB-retrieve-by-region / GB-retrieve-by-species
    bash scripts (esearch/efetch/xtract via NCBI EDirect): for every term of
    Criterion 1 (optionally crossed with Criterion 2), runs an Entrez esearch to
    get the record count, then optionally efetches the matching INSDSeq records
    (GBC/XML) to build a metadata table and/or a FASTA file.
    """

    progress     = QtCore.pyqtSignal(str)        # short status shown in the label
    log_line     = QtCore.pyqtSignal(str)         # per-query result appended to the log
    progress_pct = QtCore.pyqtSignal(int, int)    # current, total queries
    finished     = QtCore.pyqtSignal(str)         # output directory
    error        = QtCore.pyqtSignal(str)

    _NCBI_BASE   = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"
    _MAX_RETRY   = 5
    _FETCH_CHUNK = 200   # records per efetch call
    _PROGRESS_RES = 1000  # sub-steps per query, so progress also moves within
                           # a single large fetch, not only from query to query

    def __init__(self, cfg: dict, parent=None):
        super().__init__(parent)
        self.cfg   = cfg
        self._stop = False
        self._rl_lock = threading.Lock()
        self._rl_next = 0.0
        # NCBI allows ~10 req/s with an API key, ~3 req/s without — stay under both.
        self._rate = 9.0 if cfg.get("api_key") else 2.5

    def stop(self):
        self._stop = True

    # ── HTTP helpers (rate-limited, retrying — same approach as _BlastWorker) ──

    def _rate_acquire(self):
        interval = 1.0 / self._rate
        with self._rl_lock:
            now  = time.monotonic()
            wait = self._rl_next - now
            if wait > 0:
                time.sleep(wait)
            self._rl_next = time.monotonic() + interval

    def _interruptible_sleep(self, seconds: float):
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            if self._stop:
                return
            time.sleep(0.5)

    def _http_get(self, url, params, timeout=90):
        import urllib.request, urllib.parse, urllib.error
        full_url = url + "?" + urllib.parse.urlencode(params)
        for attempt in range(self._MAX_RETRY):
            if self._stop:
                return ""
            self._rate_acquire()
            if self._stop:
                return ""
            try:
                with urllib.request.urlopen(full_url, timeout=timeout) as resp:
                    if resp.status == 200:
                        return resp.read().decode("utf-8", errors="replace")
            except urllib.error.HTTPError as e:
                if e.code == 429:
                    self._interruptible_sleep(min(30 * (attempt + 1), 120))
                elif 500 <= e.code < 600:
                    self._interruptible_sleep(min(10 * (attempt + 1), 60))
                else:
                    return ""
            except Exception:
                self._interruptible_sleep(2)
        return ""

    # ── Term handling ────────────────────────────────────────────────────

    @staticmethod
    def _normalize_terms(raw_terms):
        """Strip/collapse whitespace, turn '_' into spaces, dedupe case-insensitively."""
        seen = {}
        for raw in raw_terms:
            term = re.sub(r"\s+", " ", str(raw).replace("_", " ")).strip()
            if not term:
                continue
            seen.setdefault(term.lower(), term)
        return list(seen.values())

    @staticmethod
    def _build_pairs(terms1, terms2):
        """A single term2 is a constant filter; several are crossed with every term1."""
        if not terms2:
            return [(t1, None) for t1 in terms1]
        if len(terms2) == 1:
            return [(t1, terms2[0]) for t1 in terms1]
        return [(t1, t2) for t1 in terms1 for t2 in terms2]

    # ── Entrez calls ─────────────────────────────────────────────────────

    def _esearch(self, query):
        import xml.etree.ElementTree as ET
        params = {"db": "nucleotide", "term": query, "usehistory": "y", "retmax": 0}
        if self.cfg.get("api_key"):
            params["api_key"] = self.cfg["api_key"]
        xml_text = self._http_get(f"{self._NCBI_BASE}esearch.fcgi", params)
        if not xml_text:
            return 0, "", ""
        try:
            root = ET.fromstring(xml_text)
        except Exception:
            return 0, "", ""
        count = int(root.findtext("Count", "0") or "0")
        return count, root.findtext("WebEnv", ""), root.findtext("QueryKey", "")

    @staticmethod
    def _feature_qualifiers(feat):
        """Return {qualifier_name: value} for a single INSDFeature."""
        quals = {}
        for qual in feat.findall("./INSDFeature_quals/INSDQualifier"):
            name = qual.findtext("INSDQualifier_name")
            if name and name not in quals:
                quals[name] = (qual.findtext("INSDQualifier_value") or "").strip()
        return quals

    @classmethod
    def _source_qualifiers(cls, seq_el):
        """Qualifiers of the 'source' feature (present on every INSDSeq record):
        country, isolate, specimen_voucher, collection_date, lat_lon, db_xref, etc."""
        for feat in seq_el.findall("./INSDSeq_feature-table/INSDFeature"):
            if feat.findtext("INSDFeature_key") == "source":
                return cls._feature_qualifiers(feat)
        return {}

    @classmethod
    def _record_marker(cls, seq_el):
        """Best-effort marker/gene name: first non-source feature's /gene, else /product.

        GenBank has no dedicated "marker" field like BOLD's col_Marker — this is
        free text from whichever feature (gene/CDS/rRNA/tRNA/...) the depositor
        annotated, so naming is inconsistent across records (COI/COX1/CO1/cox1...).
        """
        for feat in seq_el.findall("./INSDSeq_feature-table/INSDFeature"):
            if feat.findtext("INSDFeature_key") == "source":
                continue
            quals = cls._feature_qualifiers(feat)
            marker = quals.get("gene") or quals.get("product")
            if marker:
                return marker
        return ""

    @staticmethod
    def _safe_name(s, max_len=80):
        s = re.sub(r"[^\w\s-]", "", str(s)).strip().replace(" ", "_")
        return s[:max_len] or "unknown"

    def _iter_records(self, webenv, key, n_records, on_chunk=None):
        """Page through the esearch history set, yielding one dict per record.

        Calls `on_chunk(fetched, n_records)` after each page, so callers can
        report fetch progress for a single large query (not just query-to-query).
        """
        import xml.etree.ElementTree as ET
        fetched = 0
        while fetched < n_records:
            if self._stop:
                return
            n = min(self._FETCH_CHUNK, n_records - fetched)
            params = {
                "db": "nucleotide", "WebEnv": webenv, "query_key": key,
                "retstart": fetched, "retmax": n,
                "rettype": "gbc", "retmode": "xml",
            }
            if self.cfg.get("api_key"):
                params["api_key"] = self.cfg["api_key"]
            xml_text = self._http_get(f"{self._NCBI_BASE}efetch.fcgi", params, timeout=120)
            fetched += n
            if on_chunk:
                on_chunk(fetched, n_records)
            if not xml_text:
                continue
            try:
                root = ET.fromstring(xml_text)
            except Exception:
                continue
            for seq_el in root.findall("INSDSeq"):
                src = self._source_qualifiers(seq_el)
                taxon_id = ""
                dbxref = src.get("db_xref", "")
                if dbxref.startswith("taxon:"):
                    taxon_id = dbxref.split(":", 1)[1]

                yield {
                    "accession": (seq_el.findtext("INSDSeq_primary-accession")
                                  or seq_el.findtext("INSDSeq_locus") or ""),
                    "organism":  (seq_el.findtext("INSDSeq_organism") or "").strip(),
                    "sequence":  (seq_el.findtext("INSDSeq_sequence") or "").strip(),
                    "marker":    self._record_marker(seq_el),
                    "definition": (seq_el.findtext("INSDSeq_definition") or "").strip(),
                    "length":    seq_el.findtext("INSDSeq_length") or "",
                    "moltype":   seq_el.findtext("INSDSeq_moltype") or "",
                    "division":  (seq_el.findtext("INSDSeq_division") or "").strip(),
                    "create_date": seq_el.findtext("INSDSeq_create-date") or "",
                    "update_date": seq_el.findtext("INSDSeq_update-date") or "",
                    "country":   src.get("country", ""),
                    "isolate":   src.get("isolate", ""),
                    "specimen_voucher": src.get("specimen_voucher", ""),
                    "collection_date":  src.get("collection_date", ""),
                    "lat_lon":   src.get("lat_lon", ""),
                    "taxon_id":  taxon_id,
                    "taxonomy":  (seq_el.findtext("INSDSeq_taxonomy") or "").strip(),
                }

    @staticmethod
    def _meta_header_fmt(wb):
        return wb.add_format({"bold": True, "font_color": "FFFFFF", "bg_color": "2C4685"})

    _META_HEADERS = [
        "Accession", "Term1", "Term2", "Organism", "Marker", "Definition",
        "Length", "Moltype", "Division", "Country", "Isolate", "Specimen_Voucher",
        "Collection_Date", "Lat_Lon", "Taxon_ID", "Create_Date", "Update_Date",
        "Taxonomy",
    ]

    def _fetch_records(self, webenv, key, n_records, t1, t2, outdir,
                        ws, row, fasta_f, want_metadata, want_fasta, tally=None, on_chunk=None):
        """Fetch records for one (t1, t2) query, writing them into the compiled
        outputs (`ws`/`fasta_f`) *and* into a dedicated `outdir/<term>/` folder
        with that term's own metadata/FASTA files. If `tally` is given (a dict
        of `{column: Counter()}`), also tallies distinct values per column for
        the customizable summary report.

        Returns the next free row of the compiled worksheet.
        """
        ind_wb = ind_ws = ind_fasta_f = None
        ind_row = 1
        if want_metadata or want_fasta:
            safe = self._safe_name(t1 if not t2 else f"{t1}_{t2}")
            ind_dir = os.path.join(outdir, "individual_results", safe)
            os.makedirs(ind_dir, exist_ok=True)
            if want_metadata:
                ind_wb = xlsxwriter.Workbook(
                    os.path.join(ind_dir, f"metadata_{safe}.xlsx"), {"constant_memory": True}
                )
                ind_ws = ind_wb.add_worksheet()
                ind_ws.write_row(0, 0, self._META_HEADERS, self._meta_header_fmt(ind_wb))
            if want_fasta:
                ind_fasta_f = open(os.path.join(ind_dir, f"sequences_{safe}.fasta"), "w", encoding="utf-8")

        try:
            for rec in self._iter_records(webenv, key, n_records, on_chunk=on_chunk):
                if tally:
                    for col, counter in tally.items():
                        val = rec.get(col)
                        if val:
                            counter[val] += 1

                if want_metadata:
                    values = [
                        rec["accession"], t1, t2 or "", rec["organism"], rec["marker"],
                        rec["definition"], rec["length"], rec["moltype"], rec["division"],
                        rec["country"], rec["isolate"], rec["specimen_voucher"],
                        rec["collection_date"], rec["lat_lon"], rec["taxon_id"],
                        rec["create_date"], rec["update_date"], rec["taxonomy"],
                    ]
                    if ws is not None:
                        ws.write_row(row, 0, values)
                        row += 1
                    if ind_ws is not None:
                        ind_ws.write_row(ind_row, 0, values)
                        ind_row += 1

                if want_fasta and rec["sequence"]:
                    header_parts = [
                        rec["accession"], rec["organism"].replace(" ", "_"),
                        rec["marker"].replace(" ", "_"), t1.replace(" ", "_"),
                    ]
                    if t2:
                        header_parts.append(t2.replace(" ", "_"))
                    entry = f">{'|'.join(p for p in header_parts if p)}\n{rec['sequence'].upper()}\n"
                    if fasta_f is not None:
                        fasta_f.write(entry)
                    if ind_fasta_f is not None:
                        ind_fasta_f.write(entry)
        finally:
            if ind_wb:
                ind_wb.close()
            if ind_fasta_f:
                ind_fasta_f.close()
        return row

    # ── Customizable summary report (per-term distinct-value metrics) ──────

    _REPORT_LABELS = dict(_GENBANK_REPORT_COLUMNS)

    def _metric_headers(self, specs):
        headers = []
        for spec in specs:
            label = self._REPORT_LABELS.get(spec["column"], spec["column"])
            if spec["count"]:
                headers.append(f"# {label}")
            if spec["list"]:
                headers.append(label)
        return headers

    @staticmethod
    def _build_metric_cells(specs, tally, opts):
        cells = []
        for spec in specs:
            counter = tally.get(spec["column"])
            items = sorted(counter.items(), key=lambda kv: (-kv[1], kv[0])) if counter else []
            n_distinct = len(items)
            if spec["count"]:
                cells.append(n_distinct)
            if spec["list"]:
                max_list = opts.get("max_list", 0)
                shown = items[:max_list] if max_list > 0 else items
                if spec.get("with_counts"):
                    text = "; ".join(f"{v} ({n})" for v, n in shown)
                else:
                    text = "; ".join(v for v, n in shown)
                rest = n_distinct - len(shown)
                if rest > 0:
                    text = f"{text}; … +{rest} more" if text else f"… +{rest} more"
                cells.append(text)
        return cells

    # ── Thread entry point ──────────────────────────────────────────────

    def run(self):
        try:
            self._run_search()
        except Exception as exc:
            import traceback
            self.error.emit(f"{exc}\n{traceback.format_exc()}")

    def _run_search(self):
        cfg = self.cfg
        outdir = cfg["outdir"]
        os.makedirs(outdir, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        field1_tag = cfg["field1_tag"]
        terms1 = self._normalize_terms(cfg["field1_terms"])
        terms2 = self._normalize_terms(cfg.get("field2_terms", [])) if cfg.get("field2_enabled") else []
        field2_tag = cfg.get("field2_tag") if terms2 else None

        if not terms1:
            self.error.emit("No terms provided for Criterion 1.")
            return

        pairs = self._build_pairs(terms1, terms2)
        total = len(pairs)
        want_metadata = bool(cfg.get("want_metadata"))
        want_fasta    = bool(cfg.get("want_fasta"))
        retmax        = int(cfg.get("retmax", 500))

        metric_specs = cfg.get("report_metrics") or []
        report_opts = {
            "max_list": int(cfg.get("report_max_list", 0)) if cfg.get("report_limit_list") else 0,
        }
        want_report = bool(metric_specs)

        counts_path = os.path.join(outdir, f"genbank_batch_counts_{timestamp}.csv")
        meta_path   = os.path.join(outdir, f"genbank_batch_metadata_{timestamp}.xlsx") if want_metadata else None
        fasta_path  = os.path.join(outdir, f"genbank_batch_sequences_{timestamp}.fasta") if want_fasta else None

        counts_f = open(counts_path, "w", newline="", encoding="utf-8")
        counts_writer = csv.writer(counts_f)
        counts_writer.writerow(
            ["term1", "field1", "term2", "field2", "query", "record_count"]
            + self._metric_headers(metric_specs)
        )

        fasta_f = open(fasta_path, "w", encoding="utf-8") if fasta_path else None
        wb = ws = None
        meta_row = 1
        if meta_path:
            wb = xlsxwriter.Workbook(meta_path, {"constant_memory": True})
            ws = wb.add_worksheet()
            ws.write_row(0, 0, self._META_HEADERS, self._meta_header_fmt(wb))

        start_time = time.time()
        res = self._PROGRESS_RES
        total_units = total * res
        try:
            for i, (t1, t2) in enumerate(pairs, 1):
                if self._stop:
                    break
                query = f"{t1}[{field1_tag}]"
                if t2:
                    query += f" AND {t2}[{field2_tag}]"

                elapsed = time.time() - start_time
                rate = i / elapsed if elapsed > 0 else 0
                eta = (total - i) / rate if rate > 0 else 0
                self.progress.emit(f"({i}/{total})  {query}  —  ETA {int(eta // 60)}m {int(eta % 60)}s")
                base_units = (i - 1) * res
                self.progress_pct.emit(base_units, total_units)

                count, webenv, key = self._esearch(query)

                tally = {spec["column"]: Counter() for spec in metric_specs} if metric_specs else {}
                if count and (want_metadata or want_fasta or want_report) and webenv and key:
                    n_fetch = min(count, retmax)

                    def _on_fetch_chunk(fetched, n_records, base=base_units):
                        frac = fetched / n_records if n_records else 1.0
                        self.progress_pct.emit(base + int(frac * res), total_units)

                    meta_row = self._fetch_records(
                        webenv, key, n_fetch, t1, t2, outdir,
                        ws if want_metadata else None, meta_row,
                        fasta_f if want_fasta else None,
                        want_metadata, want_fasta,
                        tally if metric_specs else None,
                        on_chunk=_on_fetch_chunk,
                    )

                self.progress_pct.emit(i * res, total_units)

                metric_cells = self._build_metric_cells(metric_specs, tally, report_opts)
                counts_writer.writerow(
                    [t1, field1_tag, t2 or "", field2_tag or "", query, count] + metric_cells
                )
                counts_f.flush()
                self.log_line.emit(f"{query}: {count:,} record(s)")
        finally:
            counts_f.close()
            if fasta_f:
                fasta_f.close()
            if wb:
                wb.close()

        if not self._stop:
            self.log_line.emit(f"Done — {total} quer{'y' if total == 1 else 'ies'} searched.")
        self.finished.emit(outdir)


class _TermsDropEdit(QtWidgets.QPlainTextEdit):
    """Multiline term list that also accepts a dragged .txt/.csv file.

    Highlights with a hover style while a valid file is dragged over it; on
    drop, the file's lines (or first column, for .csv) replace the current text.
    """

    _ACCEPTED_EXT = (".txt", ".csv")
    _IDLE_STYLE = (
        f"QPlainTextEdit {{ background-color: #FFFFFF; "
        f"border: 1.5px dashed #C7C4BC; border-radius: 8px; padding: 6px; }}"
        f"QPlainTextEdit:focus {{ border: 1.5px dashed {BLUE_MID}; }}"
    )
    _DRAG_STYLE = (
        f"QPlainTextEdit {{ border: 2px solid {BLUE_MID}; background-color: {BLUE_LIGHT}; "
        f"border-radius: 8px; padding: 6px; }}"
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setStyleSheet(self._IDLE_STYLE)

    def _valid_url(self, url):
        return url.isLocalFile() and url.toLocalFile().lower().endswith(self._ACCEPTED_EXT)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and any(self._valid_url(u) for u in event.mimeData().urls()):
            event.acceptProposedAction()
            self.setStyleSheet(self._DRAG_STYLE)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls() and any(self._valid_url(u) for u in event.mimeData().urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self._IDLE_STYLE)
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        self.setStyleSheet(self._IDLE_STYLE)
        for url in event.mimeData().urls():
            if self._valid_url(url):
                self.load_file(url.toLocalFile())
                event.acceptProposedAction()
                return
        event.ignore()

    def load_file(self, path):
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                if path.lower().endswith(".csv"):
                    import csv as _csv
                    lines = [row[0].strip() for row in _csv.reader(f) if row and row[0].strip()]
                else:
                    lines = [ln.strip() for ln in f if ln.strip()]
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load failed", str(exc))
            return
        self.setPlainText("\n".join(lines))


class GenbankBatchPanel(QtWidgets.QWidget):
    """Batch search NCBI GenBank with up to two crossed criteria.

    Replaces the old GB-retrieve-by-region / GB-retrieve-by-species bash
    scripts: Criterion 1 is a list of terms (e.g. species names), Criterion 2
    is optional and either a single fixed value (filters every term above,
    like the old scripts' fixed `region=colombia`) or another list (crossed
    with every term above).
    """

    _FIELD_OPTIONS = [
        "Organism", "Country", "Gene", "Title",
        "All Fields", "Accession", "Author", "Journal",
    ]

    def __init__(self, parent=None):
        super().__init__(parent)

        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── Scroll area ──
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        inner = QtWidgets.QWidget()
        self._layout = QtWidgets.QVBoxLayout(inner)
        self._layout.setContentsMargins(20, 20, 20, 20)
        self._layout.setSpacing(16)
        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        # ── Title ──
        self._layout.addWidget(make_label("GenBank Batch Search", size=19, bold=True))
        desc = make_label(
            "Search NCBI GenBank (nucleotide) in batch: a list of terms for "
            "Criterion 1, optionally crossed with Criterion 2 (a single value "
            "filters every term above; a list of values is crossed with every "
            "term above — type/paste terms or drag & drop a .txt/.csv file). "
            "Produces a compiled record-count table and, optionally, a compiled "
            "metadata table and FASTA file, plus one folder per term with its own "
            "metadata/FASTA.",
            color=TEXT_SEC,
        )
        desc.setWordWrap(True)
        self._layout.addWidget(desc)

        # ── NCBI API key (optional, shared with the BLAST panel) ──
        self._layout.addWidget(self._make_section_lbl("NCBI settings"))
        api_row = QtWidgets.QHBoxLayout()
        api_row.setSpacing(8)
        api_row.addWidget(make_label("API Key (optional):", color=TEXT_SEC))
        self._api_key_edit = QtWidgets.QLineEdit()
        self._api_key_edit.setPlaceholderText("Speeds up searches (10 req/s instead of 3 req/s)…")
        self._api_key_edit.setText(self._load_api_key())
        self._api_key_edit.editingFinished.connect(self._save_api_key)
        api_row.addWidget(self._api_key_edit, 1)
        self._api_key_clear_btn = QtWidgets.QPushButton("✕")
        self._api_key_clear_btn.setFixedSize(28, 28)
        self._api_key_clear_btn.setToolTip("Clear API key")
        self._api_key_clear_btn.setStyleSheet(
            "QPushButton { background: transparent; border: 1px solid #CCC;"
            " border-radius: 4px; color: #888; font-size:11px; }"
            "QPushButton:hover { background: #FEE2E2; border-color: #EF4444; color: #EF4444; }"
        )
        self._api_key_clear_btn.clicked.connect(self._clear_api_key)
        api_row.addWidget(self._api_key_clear_btn)
        self._layout.addLayout(api_row)

        # ── Criterion 1 ──
        self._layout.addWidget(self._make_section_lbl("Criterion 1 (required)"))
        self._field1_combo, self._field1_edit, self._field1_count_lbl = self._build_criterion_row(
            "Type or paste one term per line (e.g. species names)…\n"
            "You can also drag & drop a .txt or .csv file here."
        )
        row1 = QtWidgets.QHBoxLayout()
        row1.addWidget(make_label("Field:", color=TEXT_SEC))
        row1.addWidget(self._field1_combo)
        row1.addStretch()
        row1.addWidget(self._field1_count_lbl)
        clear1_btn = QtWidgets.QPushButton("✕ Clear list")
        clear1_btn.setObjectName("secondary_btn")
        clear1_btn.setFixedWidth(120)
        clear1_btn.clicked.connect(self._field1_edit.clear)
        row1.addWidget(clear1_btn)
        self._layout.addLayout(row1)
        self._layout.addWidget(self._field1_edit)

        # ── Criterion 2 (optional) ──
        self._field2_check = QtWidgets.QCheckBox("Add second criterion (crossed with Criterion 1)")
        self._field2_check.toggled.connect(self._on_field2_toggled)
        self._layout.addWidget(self._field2_check)

        self._field2_box = QtWidgets.QWidget()
        f2_layout = QtWidgets.QVBoxLayout(self._field2_box)
        f2_layout.setContentsMargins(0, 0, 0, 0)
        f2_layout.setSpacing(6)
        self._field2_combo, self._field2_edit, self._field2_count_lbl = self._build_criterion_row(
            "One value applies to all terms above; several values are crossed…\n"
            "You can also drag & drop a .txt or .csv file here."
        )
        row2 = QtWidgets.QHBoxLayout()
        row2.addWidget(make_label("Field:", color=TEXT_SEC))
        row2.addWidget(self._field2_combo)
        row2.addStretch()
        row2.addWidget(self._field2_count_lbl)
        clear2_btn = QtWidgets.QPushButton("✕ Clear list")
        clear2_btn.setObjectName("secondary_btn")
        clear2_btn.setFixedWidth(120)
        clear2_btn.clicked.connect(self._field2_edit.clear)
        row2.addWidget(clear2_btn)
        f2_layout.addLayout(row2)
        f2_layout.addWidget(self._field2_edit)
        self._field2_box.hide()
        self._layout.addWidget(self._field2_box)

        # ── Options ──
        self._layout.addWidget(self._make_section_lbl("Options"))

        cap_row = QtWidgets.QHBoxLayout()
        cap_row.setSpacing(8)
        cap_row.addWidget(make_label("Records per query (cap):", color=TEXT_SEC))
        self._retmax_spin = QtWidgets.QSpinBox()
        self._retmax_spin.setRange(1, 10000)
        self._retmax_spin.setValue(500)
        self._retmax_spin.setFixedWidth(100)
        cap_row.addWidget(self._retmax_spin)
        warn_icon = QtWidgets.QLabel("⚠")
        warn_icon.setStyleSheet("color:#B45309; font-size:17px; padding:0 2px;")
        warn_icon.setToolTip(
            "<b>NCBI usage policy warning</b><br><br>"
            "NCBI monitors and penalizes excessive use of its servers.<br>"
            "A high cap combined with long term lists can generate many "
            "requests — keep batches reasonable and avoid running several "
            "sessions in parallel."
        )
        cap_row.addWidget(warn_icon)
        cap_row.addStretch()
        self._layout.addLayout(cap_row)

        self._meta_check  = QtWidgets.QCheckBox("Metadata (XLSX)")
        self._fasta_check = QtWidgets.QCheckBox("Sequences (FASTA)")
        counts_check = QtWidgets.QCheckBox("Record counts (CSV)")
        counts_check.setChecked(True)
        counts_check.setEnabled(False)
        out_row = QtWidgets.QHBoxLayout()
        out_row.addWidget(counts_check)
        out_row.addWidget(self._meta_check)
        out_row.addWidget(self._fasta_check)
        out_row.addStretch()
        self._layout.addLayout(out_row)

        # ── Customize summary report (optional) ──
        self._report_check = QtWidgets.QCheckBox("Customize summary report (optional)")
        self._report_check.toggled.connect(self._on_report_toggled)
        self._layout.addWidget(self._report_check)

        self._report_box = QtWidgets.QWidget()
        report_layout = QtWidgets.QVBoxLayout(self._report_box)
        report_layout.setContentsMargins(0, 0, 0, 0)
        report_layout.setSpacing(8)

        report_hint = make_label(
            "Adds extra columns to the counts CSV: for each column below, how "
            "many distinct values were found per term, and/or which ones. "
            "Requires fetching each record from NCBI (subject to the cap above), "
            "even if metadata/FASTA output isn't saved.",
            size=18, color=TEXT_HINT,
        )
        report_hint.setWordWrap(True)
        report_layout.addWidget(report_hint)

        grid = QtWidgets.QGridLayout()
        grid.setSpacing(6)
        grid.addWidget(make_label("Column", color=TEXT_SEC, bold=True), 0, 0)
        grid.addWidget(make_label("Count", color=TEXT_SEC, bold=True), 0, 1,
                        QtCore.Qt.AlignCenter)
        grid.addWidget(make_label("List values", color=TEXT_SEC, bold=True), 0, 2,
                        QtCore.Qt.AlignCenter)
        grid.addWidget(make_label("List values counts", color=TEXT_SEC, bold=True), 0, 3,
                        QtCore.Qt.AlignCenter)
        self._report_checks: dict = {}
        for i, (col_key, label) in enumerate(_GENBANK_REPORT_COLUMNS, start=1):
            grid.addWidget(make_label(label, color=TEXT_SEC), i, 0)
            count_chk  = QtWidgets.QCheckBox()
            list_chk   = QtWidgets.QCheckBox()
            counts_chk = QtWidgets.QCheckBox()
            counts_chk.setToolTip("Show how many times each value appears, e.g. “Colombia (12)”")
            grid.addWidget(count_chk, i, 1, QtCore.Qt.AlignCenter)
            grid.addWidget(list_chk, i, 2, QtCore.Qt.AlignCenter)
            grid.addWidget(counts_chk, i, 3, QtCore.Qt.AlignCenter)
            self._report_checks[col_key] = (count_chk, list_chk, counts_chk)
        report_layout.addLayout(grid)

        report_opts_row = QtWidgets.QHBoxLayout()
        report_opts_row.setSpacing(8)
        self._report_limit_chk = QtWidgets.QCheckBox("Limit list to")
        report_opts_row.addWidget(self._report_limit_chk)
        self._report_limit_spin = QtWidgets.QSpinBox()
        self._report_limit_spin.setRange(1, 1000)
        self._report_limit_spin.setValue(10)
        self._report_limit_spin.setFixedWidth(70)
        self._report_limit_spin.setEnabled(False)
        self._report_limit_chk.toggled.connect(self._report_limit_spin.setEnabled)
        report_opts_row.addWidget(self._report_limit_spin)
        report_opts_row.addWidget(make_label("values", color=TEXT_SEC))
        report_opts_row.addStretch()
        report_reset_btn = QtWidgets.QPushButton("Reset")
        report_reset_btn.setObjectName("secondary_btn")
        report_reset_btn.setFixedWidth(90)
        report_reset_btn.clicked.connect(self._reset_report_spec)
        report_opts_row.addWidget(report_reset_btn)
        report_layout.addLayout(report_opts_row)

        self._report_box.hide()
        self._layout.addWidget(self._report_box)

        self._layout.addStretch()

        # ── Status + progress (pinned outside the scroll area, so a tall
        #    scrollable section above — e.g. the report grid — never pushes
        #    them out of view) ──
        status_bar = QtWidgets.QWidget()
        status_bar.setObjectName("genbank_batch_status_bar")
        status_bar.setStyleSheet(f"""
            QWidget#genbank_batch_status_bar {{
                background: {GRAY_CARD};
                border-top: 1px solid {GRAY_LINE};
            }}
        """)
        status_layout = QtWidgets.QVBoxLayout(status_bar)
        status_layout.setContentsMargins(20, 10, 20, 10)
        status_layout.setSpacing(6)

        self._status_lbl = make_label("", color=TEXT_SEC)
        status_layout.addWidget(self._status_lbl)

        self._progress_bar = QtWidgets.QProgressBar()
        self._progress_bar.setFixedHeight(22)
        self._progress_bar.setTextVisible(True)
        self._progress_bar.setFormat("%p%")
        self._progress_bar.setAlignment(QtCore.Qt.AlignCenter)
        self._progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 1px solid {GRAY_LINE};
                border-radius: 6px;
                background: {GRAY_BG};
                color: {TEXT_PRI};
                font-size: 15px;
                font-weight: 600;
                text-align: center;
            }}
            QProgressBar::chunk {{
                background-color: {BLUE};
                border-radius: 5px;
            }}
        """)
        self._progress_bar.hide()
        status_layout.addWidget(self._progress_bar)

        outer.addWidget(status_bar)

        # ── Log ──
        self._log_edit = QtWidgets.QPlainTextEdit()
        self._log_edit.setReadOnly(True)
        self._log_edit.setFixedHeight(160)
        self._log_edit.setStyleSheet(f"""
            QPlainTextEdit {{
                background: {GRAY_BG};
                border: 1px solid {GRAY_LINE};
                border-top: 1px solid {GRAY_LINE};
                border-radius: 0px;
                font-family: 'Courier New', Consolas, monospace;
                font-size: 17px;
                color: {TEXT_PRI};
                padding: 6px;
            }}
        """)
        self._log_edit.hide()
        outer.addWidget(self._log_edit)

        # ── Footer ──
        footer = QtWidgets.QWidget()
        footer.setObjectName("genbank_batch_footer")
        footer.setStyleSheet(f"""
            QWidget#genbank_batch_footer {{
                background: {GRAY_CARD};
                border-top: 1px solid {GRAY_LINE};
            }}
        """)
        fl = QtWidgets.QHBoxLayout(footer)
        fl.setContentsMargins(20, 10, 20, 10)
        fl.setSpacing(8)

        self._clear_btn = QtWidgets.QPushButton("Clear")
        self._clear_btn.setObjectName("danger_btn")
        self._clear_btn.setFixedHeight(44)
        self._clear_btn.setFixedWidth(140)
        self._clear_btn.clicked.connect(self._clear)
        fl.addWidget(self._clear_btn)

        self._open_folder_btn = QtWidgets.QPushButton("Open folder  📂")
        self._open_folder_btn.setObjectName("secondary_btn")
        self._open_folder_btn.setFixedHeight(44)
        self._open_folder_btn.hide()
        self._open_folder_btn.clicked.connect(self._open_output_folder)
        fl.addWidget(self._open_folder_btn)

        self._stop_btn = QtWidgets.QPushButton("Stop")
        self._stop_btn.setObjectName("danger_btn")
        self._stop_btn.setFixedHeight(44)
        self._stop_btn.setFixedWidth(120)
        self._stop_btn.hide()
        self._stop_btn.clicked.connect(self._stop_search)
        fl.addWidget(self._stop_btn)

        fl.addStretch()

        self._run_btn = QtWidgets.QPushButton("Search GenBank  →")
        self._run_btn.setObjectName("primary_btn")
        self._run_btn.setFixedHeight(44)
        self._run_btn.setFixedWidth(300)
        self._run_btn.clicked.connect(self._on_run_clicked)
        fl.addWidget(self._run_btn)
        self._set_run_enabled(False)
        outer.addWidget(footer)

        self._field1_edit.textChanged.connect(
            lambda: self._set_run_enabled(bool(self._field1_edit.toPlainText().strip()))
        )

        self._worker: Optional[_GenbankBatchWorker] = None
        self._retired_workers: set = set()
        self._last_outdir = ""

    # ── UI helpers ────────────────────────────────────────────────────────

    @staticmethod
    def _make_section_lbl(text):
        lbl = make_section_label(text)
        lbl.setStyleSheet(f"font-size:17px; font-weight:600; color:{BLUE}; letter-spacing:0.4px;")
        return lbl

    def _build_criterion_row(self, placeholder):
        combo = QtWidgets.QComboBox()
        combo.addItems(self._FIELD_OPTIONS)
        combo.setFixedWidth(200)
        edit = _TermsDropEdit()
        edit.setPlaceholderText(placeholder)
        edit.setFixedHeight(90)
        count_lbl = make_label("0 terms", size=18, bold=True, color=BLUE)
        edit.textChanged.connect(lambda: self._update_term_count(edit, count_lbl))
        return combo, edit, count_lbl

    @staticmethod
    def _update_term_count(edit: QtWidgets.QPlainTextEdit, label: QtWidgets.QLabel):
        n = len([ln for ln in edit.toPlainText().splitlines() if ln.strip()])
        label.setText(f"{n} term{'s' if n != 1 else ''}")

    def _on_field2_toggled(self, checked):
        self._field2_box.setVisible(checked)

    def _on_report_toggled(self, checked):
        self._report_box.setVisible(checked)

    def _reset_report_spec(self):
        for count_chk, list_chk, counts_chk in self._report_checks.values():
            count_chk.setChecked(False)
            list_chk.setChecked(False)
            counts_chk.setChecked(False)
        self._report_limit_chk.setChecked(False)
        self._report_limit_spin.setValue(10)

    def _collect_report_spec(self):
        metrics = []
        for col_key, (count_chk, list_chk, counts_chk) in self._report_checks.items():
            if count_chk.isChecked() or list_chk.isChecked():
                metrics.append({
                    "column": col_key,
                    "count": count_chk.isChecked(),
                    "list":  list_chk.isChecked(),
                    "with_counts": list_chk.isChecked() and counts_chk.isChecked(),
                })
        return metrics

    def _set_run_enabled(self, enabled: bool):
        self._run_btn.setEnabled(enabled)
        if enabled:
            self._run_btn.setStyleSheet(
                f"QPushButton {{ background-color:{BLUE}; color:white; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
                f"QPushButton:hover {{ background-color:#0C4A82; }}"
            )
        else:
            self._run_btn.setStyleSheet(
                f"QPushButton {{ background-color:{GRAY_LINE}; color:{TEXT_HINT}; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
            )

    # ── API key persistence (shared file with BlastPanel) ──────────────────

    @staticmethod
    def _config_path():
        return os.path.join(_profiles_dir(), "blast_config.json")

    def _load_api_key(self) -> str:
        try:
            with open(self._config_path(), "r", encoding="utf-8") as f:
                return _json_mod.load(f).get("api_key", "")
        except Exception:
            return ""

    def _save_api_key(self):
        path = self._config_path()
        try:
            data: dict = {}
            if os.path.isfile(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        data = _json_mod.load(f)
                except Exception:
                    pass
            data["api_key"] = self._api_key_edit.text().strip()
            with open(path, "w", encoding="utf-8") as f:
                _json_mod.dump(data, f, indent=2)
        except Exception:
            pass

    def _clear_api_key(self):
        self._api_key_edit.clear()
        self._save_api_key()

    # ── Output folder dialog (same pattern as BlastPanel._ask_output_dir) ──

    def _ask_output_dir(self, auto_dir: str, folder_name: str) -> Optional[str]:
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Output folder")
        dlg.setMinimumWidth(480)
        dlg.setStyleSheet(f"""
            QDialog {{ background-color: {GRAY_CARD}; }}
            QLabel {{ color: {TEXT_PRI}; background-color: transparent; }}
            QRadioButton {{
                color: {TEXT_PRI}; background-color: transparent;
                font-size: 15px; padding: 6px 0;
            }}
            QRadioButton::indicator {{ width: 16px; height: 16px; }}
            QPushButton {{
                border-radius: 8px; padding: 8px 20px;
                font-size: 15px; font-weight: 500;
            }}
            #dlg_ok_btn {{ background-color: {BLUE}; color: white; border: none; }}
            #dlg_ok_btn:hover {{ background-color: #0C4A82; }}
            #dlg_cancel_btn {{
                background-color: transparent; color: {BLUE};
                border: 1px solid {BLUE};
            }}
            #dlg_cancel_btn:hover {{ background-color: {BLUE_LIGHT}; }}
        """)

        vlay = QtWidgets.QVBoxLayout(dlg)
        vlay.setSpacing(16)
        vlay.setContentsMargins(24, 24, 24, 20)

        title_lbl = QtWidgets.QLabel("Where to save the results?")
        title_lbl.setStyleSheet(f"font-size:17px; font-weight:700; color:{TEXT_PRI};")
        vlay.addWidget(title_lbl)

        radio_auto = QtWidgets.QRadioButton(
            f"Automatic folder (recommended)\n  …/output/{folder_name}/"
        )
        radio_auto.setChecked(True)
        radio_custom = QtWidgets.QRadioButton("Select folder manually")
        vlay.addWidget(radio_auto)
        vlay.addWidget(radio_custom)
        vlay.addSpacing(8)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_cancel.setObjectName("dlg_cancel_btn")
        btn_cancel.setFixedHeight(38)
        btn_ok = QtWidgets.QPushButton("Run")
        btn_ok.setObjectName("dlg_ok_btn")
        btn_ok.setFixedHeight(38)
        btn_ok.setDefault(True)
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addSpacing(8)
        btn_row.addWidget(btn_ok)
        vlay.addLayout(btn_row)

        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return None

        if radio_auto.isChecked():
            return auto_dir
        parent_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Select output folder")
        if not parent_dir:
            return None
        return os.path.join(parent_dir, folder_name)

    # ── Run lifecycle ────────────────────────────────────────────────────

    def _on_run_clicked(self):
        field1_terms = [ln for ln in self._field1_edit.toPlainText().splitlines() if ln.strip()]
        if not field1_terms:
            QtWidgets.QMessageBox.warning(self, "No terms", "Enter at least one term for Criterion 1.")
            return

        field2_enabled = self._field2_check.isChecked()
        field2_terms = (
            [ln for ln in self._field2_edit.toPlainText().splitlines() if ln.strip()]
            if field2_enabled else []
        )

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"genbank_batch_{ts}"
        auto_dir = os.path.join(_get_base_dir(), "output", folder_name)
        out_dir = self._ask_output_dir(auto_dir, folder_name)
        if out_dir is None:
            return

        self._save_api_key()

        cfg = {
            "api_key":        self._api_key_edit.text().strip(),
            "field1_tag":     self._field1_combo.currentText(),
            "field1_terms":   field1_terms,
            "field2_enabled": field2_enabled,
            "field2_tag":     self._field2_combo.currentText(),
            "field2_terms":   field2_terms,
            "retmax":         self._retmax_spin.value(),
            "want_metadata":  self._meta_check.isChecked(),
            "want_fasta":     self._fasta_check.isChecked(),
            "report_metrics":     self._collect_report_spec() if self._report_check.isChecked() else [],
            "report_limit_list":  self._report_limit_chk.isChecked(),
            "report_max_list":    self._report_limit_spin.value(),
            "outdir":         out_dir,
        }

        total_pairs = len(_GenbankBatchWorker._build_pairs(
            _GenbankBatchWorker._normalize_terms(field1_terms),
            _GenbankBatchWorker._normalize_terms(field2_terms) if field2_enabled else [],
        ))

        self._retire_worker()
        self._log_edit.clear()
        self._log_edit.show()
        self._progress_bar.setRange(0, max(1, total_pairs * _GenbankBatchWorker._PROGRESS_RES))
        self._progress_bar.setValue(0)
        self._progress_bar.show()
        self._status_lbl.setStyleSheet(f"color:{TEXT_SEC};")
        self._status_lbl.setText(f"Starting search — {total_pairs} quer{'y' if total_pairs == 1 else 'ies'}…")

        self._worker = _GenbankBatchWorker(cfg)
        self._worker.progress.connect(self._on_progress)
        self._worker.log_line.connect(self._log_edit.appendPlainText)
        self._worker.progress_pct.connect(self._on_progress_pct)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._set_running(True)
        self._worker.start()

    def _on_progress(self, text: str):
        self._status_lbl.setText(text)

    def _on_progress_pct(self, current: int, total: int):
        if total > 0 and self._progress_bar.maximum() != total:
            self._progress_bar.setRange(0, total)
        self._progress_bar.setValue(current)

    def _set_running(self, running: bool):
        self._run_btn.setVisible(not running)
        self._stop_btn.setVisible(running)
        self._clear_btn.setEnabled(not running)

    def _stop_search(self):
        if self._worker and self._worker.isRunning():
            self._worker.stop()
            self._status_lbl.setText("Stopping… (finishing current request)")

    def _retire_worker(self):
        """Detach the current worker so a late signal can't touch the UI, and keep a
        reference until its thread actually exits (a QThread garbage-collected while
        still running crashes Qt)."""
        self._retired_workers = {w for w in self._retired_workers if w.isRunning()}
        w = self._worker
        self._worker = None
        if w is None:
            return
        for sig in (w.progress, w.log_line, w.progress_pct, w.finished, w.error):
            try:
                sig.disconnect()
            except (TypeError, RuntimeError):
                pass
        if w.isRunning():
            w.stop()
            self._retired_workers.add(w)

    def _on_finished(self, outdir: str):
        self._set_running(False)
        self._progress_bar.hide()
        self._last_outdir = outdir
        self._status_lbl.setStyleSheet(f"color:{GREEN};")
        self._status_lbl.setText(f"Done. Results saved in: {outdir}")
        if outdir and os.path.isdir(outdir):
            self._open_folder_btn.show()

    def _on_error(self, msg: str):
        self._set_running(False)
        self._progress_bar.hide()
        self._status_lbl.setStyleSheet(f"color:{RED};")
        self._status_lbl.setText(f"Error: {msg}")
        self._log_edit.show()
        self._log_edit.appendPlainText(f"ERROR: {msg}")

    def _open_output_folder(self):
        if self._last_outdir and os.path.isdir(self._last_outdir):
            os.startfile(self._last_outdir)

    def _clear(self):
        self._field1_edit.clear()
        self._field2_check.setChecked(False)
        self._field2_edit.clear()
        self._meta_check.setChecked(False)
        self._fasta_check.setChecked(False)
        self._retmax_spin.setValue(500)
        self._report_check.setChecked(False)
        self._reset_report_spec()
        self._log_edit.clear()
        self._log_edit.hide()
        self._progress_bar.hide()
        self._status_lbl.setText("")
        self._open_folder_btn.hide()
        self._last_outdir = ""
        self._set_run_enabled(False)


# ═══════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ═══════════════════════════════════════════════════════════════════════════


class _BlastWorker(QtCore.QThread):
    """
    Python equivalent of blast-tax-remote.sh.
    Submits multiFASTA batches to NCBI BLAST via the CGI API, then retrieves
    organism names (nucleotide DB) and taxonomic classification (taxonomy DB)
    for each hit accession.  Results are saved as a TSV file, written per
    batch so data is never lost if the run is interrupted.
    """
    statusUpdated   = QtCore.pyqtSignal(str, str)   # (slot_key, text)
    progressUpdated = QtCore.pyqtSignal(int, int)   # current, total
    taskFinished    = QtCore.pyqtSignal(str)         # output directory
    taskError       = QtCore.pyqtSignal(str)

    _NCBI_BASE    = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"
    _BLAST_URL    = "https://blast.ncbi.nlm.nih.gov/blast/Blast.cgi"
    _MAX_RETRY    = 50
    _POLL_SLEEP   = 5    # seconds between BLAST status polls
    _SOCK_TIMEOUT = 15   # max seconds blocked in a single urlopen call
    _TAX_RETRIES  = 2    # extra retry rounds for "Not_found_in_Taxonomy" results
    _NCBI_RATE    = 9.0  # max HTTP requests/second (NCBI allows 10 with API key)
    _BATCH_UNITS  = 1000 # progress units per batch (phases: wait 0-400, org 400-600, tax 600-800, rows 800-1000)

    def __init__(self, files: List[str], cfg: dict, parent=None):
        super().__init__(parent)
        self.files  = files
        self.cfg    = cfg
        self._stop  = False

        # Shared cache dicts (populated in _run_blast, accessed from threads)
        self._accdb:  dict = {}
        self._taxadb: dict = {}
        self._cache_lock = threading.Lock()
        # Keys already persisted to disk — only new ones are appended per batch
        self._saved_acc_keys:  set = set()
        self._saved_tax_keys:  set = set()
        # Org keys whose CURRENT _taxadb value is a negative that came from a
        # FAILED efetch (transient/network), not from a real "no lineage" answer.
        # These are kept in memory (so the in-run retry loop still finds them) but
        # must NOT be persisted to taxadb.dbx — otherwise a network blip would mark
        # an organism "Not_found_in_Taxonomy" permanently across future runs.
        self._tax_unconfirmed: set = set()

        # Internal map built during batch organism fetch: org_key → taxid string
        # Used to skip the esearch step when fetching taxonomy
        self._org_to_taxid: dict = {}

        # Rate limiter: serialises HTTP calls across all worker threads so
        # we stay within NCBI's 10 req/s limit (using 9 for safety margin).
        self._rl_lock     = threading.Lock()
        self._rl_next     = 0.0   # monotonic time of next allowed request

        # Usage monitor — accessed from multiple threads via _usage_lock
        self._usage_lock       = threading.Lock()
        self._req_count        = 0
        self._rate_limit_count = 0
        self._server_err_count = 0

    def stop(self):
        self._stop = True

    def _rate_acquire(self):
        """Block until the next NCBI request slot is available."""
        interval = 1.0 / self._NCBI_RATE
        with self._rl_lock:
            now  = time.monotonic()
            wait = self._rl_next - now
            if wait > 0:
                time.sleep(wait)
            self._rl_next = time.monotonic() + interval

    # ── HTTP helpers ──────────────────────────────────────────────────────

    def _http_get(self, url, params=None, timeout=60):
        import urllib.request, urllib.parse, urllib.error
        if params:
            url = url + "?" + urllib.parse.urlencode(params)
        for attempt in range(self._MAX_RETRY):
            if self._stop:
                return ""
            self._rate_acquire()
            if self._stop:
                return ""
            try:
                req = urllib.request.Request(url)
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    if resp.status == 200:
                        text = resp.read().decode("utf-8", errors="replace")
                        if text:
                            with self._usage_lock:
                                self._req_count += 1
                            return text
            except urllib.error.HTTPError as e:
                if e.code == 429:
                    with self._usage_lock:
                        self._rate_limit_count += 1
                    try:
                        wait = int(e.headers.get("Retry-After") or 0)
                    except Exception:
                        wait = 0
                    if wait <= 0:
                        wait = min(30 * (attempt + 1), 120)
                    self._interruptible_sleep(wait)
                elif 500 <= e.code < 600:
                    with self._usage_lock:
                        self._server_err_count += 1
                    self._interruptible_sleep(min(10 * (attempt + 1), 60))
                else:
                    return ""  # 4xx other than 429 — not retryable
            except Exception:
                self._interruptible_sleep(2)
        return ""

    def _http_post(self, url, data: str, timeout=120):
        import urllib.request, urllib.error
        for attempt in range(self._MAX_RETRY):
            if self._stop:
                return ""
            self._rate_acquire()
            if self._stop:
                return ""
            try:
                req = urllib.request.Request(
                    url,
                    data=data.encode("utf-8"),
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                    method="POST"
                )
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    if resp.status == 200:
                        text = resp.read().decode("utf-8", errors="replace")
                        if text:
                            with self._usage_lock:
                                self._req_count += 1
                            return text
            except urllib.error.HTTPError as e:
                if e.code == 429:
                    with self._usage_lock:
                        self._rate_limit_count += 1
                    try:
                        wait = int(e.headers.get("Retry-After") or 0)
                    except Exception:
                        wait = 0
                    if wait <= 0:
                        wait = min(30 * (attempt + 1), 120)
                    self._interruptible_sleep(wait)
                elif 500 <= e.code < 600:
                    with self._usage_lock:
                        self._server_err_count += 1
                    self._interruptible_sleep(min(10 * (attempt + 1), 60))
                else:
                    return ""  # 4xx other than 429 — not retryable
            except Exception:
                self._interruptible_sleep(2)
        return ""

    def _interruptible_sleep(self, seconds: float):
        """Sleep in 0.5 s chunks so _stop is checked frequently."""
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            if self._stop:
                return
            time.sleep(0.5)

    # ── FASTA helpers ─────────────────────────────────────────────────────

    def _merge_fasta_files(self, files):
        parts = []
        for f in files:
            try:
                with open(f, "r", encoding="utf-8", errors="replace") as fh:
                    content = fh.read().replace("\r\n", "\n").replace("\r", "\n")
                    parts.append(content.strip())
            except Exception as e:
                self.statusUpdated.emit("blast", f"BLAST       │ Warning: could not read {os.path.basename(f)}")
        return "\n".join(parts)

    def _to_single_line_fasta(self, text):
        out = []
        header = None
        seq_parts = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith(">"):
                if header is not None:
                    out.append(header)
                    out.append("".join(seq_parts))
                header = line.replace(" ", "_")
                seq_parts = []
            else:
                seq_parts.append(line)
        if header is not None:
            out.append(header)
            out.append("".join(seq_parts))
        return "\n".join(out)

    def _split_batches(self, fasta_text, nseq):
        """Return (batches, batch_pairs_list).
        batches[i]       – FASTA text string for batch i
        batch_pairs_list[i] – list of (header, seq) tuples for batch i
        """
        # `fasta_text` comes from _to_single_line_fasta(): each header is followed by
        # exactly one sequence line (possibly empty). Do NOT drop blank lines here —
        # filtering an empty sequence line would shift the next header into its place
        # and desync every header/seq pair from that point on.
        lines = fasta_text.splitlines()
        pairs = []
        i = 0
        while i < len(lines):
            if lines[i].startswith(">"):
                header = lines[i]
                seq    = lines[i + 1] if i + 1 < len(lines) else ""
                pairs.append((header, seq))
                i += 2
            else:
                i += 1
        batches = []
        batch_pairs_list = []
        for start in range(0, len(pairs), nseq):
            chunk = pairs[start:start + nseq]
            batches.append("\n".join(h + "\n" + s for h, s in chunk))
            batch_pairs_list.append(chunk)
        return batches, batch_pairs_list

    # ── BLAST API ─────────────────────────────────────────────────────────

    def _url_encode_fasta(self, text):
        from urllib.parse import quote
        safe = (
            "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            "abcdefghijklmnopqrstuvwxyz"
            "0123456789_-"
        )
        return quote(text, safe=safe)

    def _blast_submit(self, fasta_text):
        cfg = self.cfg
        encoded = self._url_encode_fasta(fasta_text)
        data = (
            f"CMD=Put&PROGRAM={cfg['program']}&DATABASE={cfg['database']}"
            f"&api_key={cfg['api_key']}&HITLIST_SIZE={cfg['nhits']}&QUERY={encoded}"
        )
        resp = self._http_post(self._BLAST_URL, data)
        if not resp:
            return None, 30
        import re
        rid_m  = re.search(r"RID = ([^\s]+)\s+RTOE", resp)
        rtoe_m = re.search(r"RTOE = (\d+)", resp)
        rid  = rid_m.group(1) if rid_m else None
        rtoe = int(rtoe_m.group(1)) if rtoe_m else 30
        return rid, rtoe

    def _blast_poll(self, rid, batch_label=""):
        url = f"{self._BLAST_URL}?CMD=Get&RID={rid}"
        attempts = 0
        t0 = time.monotonic()
        prefix = f"BLAST       │ [{batch_label}] " if batch_label else "BLAST       │ "
        while not self._stop and attempts < self._MAX_RETRY:
            resp = self._http_get(url)
            if self._stop:
                return False
            if "Status=WAITING" in resp:
                elapsed = int(time.monotonic() - t0)
                self.statusUpdated.emit(
                    "blast",
                    f"{prefix}Polling for results… "
                    f"(attempt {attempts + 1}/{self._MAX_RETRY} · {elapsed}s elapsed)"
                )
                self._interruptible_sleep(self._POLL_SLEEP)
                attempts += 1
                continue
            if "Status=FAILED" in resp:
                self.statusUpdated.emit("blast", f"{prefix}Search failed for RID {rid}.")
                return False
            if "Status=UNKNOWN" in resp:
                self.statusUpdated.emit("blast", f"{prefix}Search expired for RID {rid}.")
                return False
            if "Status=READY" in resp:
                return True
            # Empty or unrecognised response — count as a transient failure
            elapsed = int(time.monotonic() - t0)
            self.statusUpdated.emit(
                "blast",
                f"{prefix}Polling for results… "
                f"(attempt {attempts + 1}/{self._MAX_RETRY} · {elapsed}s elapsed)"
            )
            attempts += 1
            self._interruptible_sleep(self._POLL_SLEEP)
        if not self._stop:
            elapsed = int(time.monotonic() - t0)
            self.statusUpdated.emit(
                "blast",
                f"{prefix}No response after {attempts} polls ({elapsed}s)."
            )
        return False

    def _blast_get_tabular(self, rid):
        url = (
            f"{self._BLAST_URL}"
            f"?CMD=Get&FORMAT_TYPE=Text&ALIGNMENT_VIEW=Tabular&RID={rid}"
        )
        resp = self._http_get(url, timeout=120)
        return self._parse_tabular(resp)

    def _parse_tabular(self, text):
        # Equivalent to: grep -A nhits '^#' | sed '/^#/d; /^--/d'
        rows = []
        nhits   = self.cfg["nhits"]
        capture = False
        count   = 0
        for line in text.splitlines():
            if line.startswith("# Query:") or (
                line.startswith("#") and "Fields:" in line
            ):
                capture = True
                count   = 0
                continue
            if capture:
                if line.startswith("#") or line.startswith("--"):
                    continue
                stripped = line.strip()
                if stripped:
                    rows.append(stripped)
                    count += 1
                    if count >= nhits:
                        capture = False
        return rows

    @staticmethod
    def _rank_rows(rows):
        """Prepend a per-sample Hit_rank column (1 = best hit) to each tabular row.

        `rows` must arrive in BLAST quality order (best first) within each query,
        as produced by _parse_tabular — so the rank is assigned BEFORE any
        alphabetical reordering. The output is grouped by Query_name with hits in
        ascending rank, and each row gains the rank as its first tab field."""
        counter = {}
        ranked = []
        for row in rows:
            query = row.split("\t", 1)[0]
            rank = counter.get(query, 0) + 1
            counter[query] = rank
            ranked.append((query, rank, row))
        ranked.sort(key=lambda t: (t[0], t[1]))
        return [f"{rank}\t{row}" for (query, rank, row) in ranked]

    # ── NCBI metadata ────────────────────────────────────────────────────

    def _fetch_organism(self, accession: str) -> str:
        """Return organism name for *accession*, using in-memory cache."""
        with self._cache_lock:
            if accession in self._accdb:
                return self._accdb[accession]
        if self._stop:
            return ""
        import xml.etree.ElementTree as ET
        base = self._NCBI_BASE
        xml = self._http_get(
            f"{base}esearch.fcgi",
            params={"db": "nucleotide", "term": accession, "usehistory": "y"}
        )
        if not xml or self._stop:
            return ""   # transient failure — do NOT cache so a later batch can retry
        try:
            root = ET.fromstring(xml)
            web = root.findtext("WebEnv", "")
            key = root.findtext("QueryKey", "")
        except Exception:
            return ""
        if not web or not key:
            return ""
        xml2 = self._http_get(
            f"{base}efetch.fcgi",
            params={
                "db": "nucleotide", "query_key": key,
                "WebEnv": web, "rettype": "gbc", "retmode": "xml"
            },
            timeout=90
        )
        if not xml2 or self._stop:
            return ""   # transient failure — do NOT cache
        try:
            root2 = ET.fromstring(xml2)
            el = root2.find(".//INSDSeq_organism")
            if el is None or not el.text:
                return ""
            org = (
                el.text.strip()
                .replace(" ", "_")
                .replace("&apos;", "'")
                .replace("&amp;", "&")
            )
        except Exception:
            return ""
        with self._cache_lock:
            self._accdb[accession] = org
        return org

    def _fetch_taxonomy(self, organism: str) -> str:
        """Return taxonomy string for *organism*, using in-memory cache."""
        with self._cache_lock:
            if organism in self._taxadb:
                return self._taxadb[organism]
        if self._stop:
            return "Not_found_in_Taxonomy"
        import xml.etree.ElementTree as ET
        base = self._NCBI_BASE
        # NCBI scientific names use spaces; organism is stored with underscores
        search_name = organism.replace("_", " ")
        xml = self._http_get(
            f"{base}esearch.fcgi",
            params={
                "db": "taxonomy",
                "term": f'"{search_name}"[Scientific Name]',
                "usehistory": "y"
            }
        )
        if not xml or self._stop:
            return "Not_found_in_Taxonomy"   # transient — do NOT cache
        try:
            root = ET.fromstring(xml)
            web = root.findtext("WebEnv", "")
            key = root.findtext("QueryKey", "")
        except Exception:
            return "Not_found_in_Taxonomy"
        if not web or not key:
            return "Not_found_in_Taxonomy"
        xml2 = self._http_get(
            f"{base}efetch.fcgi",
            params={
                "db": "taxonomy", "query_key": key,
                "WebEnv": web, "retmode": "xml"
            },
            timeout=90
        )
        if not xml2 or self._stop:
            return "Not_found_in_Taxonomy"   # transient — do NOT cache
        tax = self._parse_taxonomy_xml(xml2)
        result = tax if tax else "Not_found_in_Taxonomy"
        # Cache confirmed results (both found and genuinely absent)
        with self._cache_lock:
            self._taxadb[organism] = result
        return result

    def _parse_taxonomy_xml(self, xml_text):
        import xml.etree.ElementTree as ET
        try:
            root = ET.fromstring(xml_text)
        except Exception:
            return ""
        ranks_wanted = {"kingdom", "class", "order", "family", "genus"}
        found = {}
        for taxon in root.iter("Taxon"):
            rank = (taxon.findtext("Rank") or "").strip().lower()
            name = (taxon.findtext("ScientificName") or "").strip()
            if rank in ranks_wanted and rank not in found:
                found[rank] = name
        return "\t".join(
            found.get(r, "-") for r in ["kingdom", "class", "order", "family", "genus"]
        )

    # ── Batch fetch methods ───────────────────────────────────────────────

    def _fetch_organisms_batch(self, accessions: List[str]) -> None:
        """Fetch organism names + taxids for all accessions in one POST per 500-ID chunk.

        Populates _accdb (accession → org_key) and _org_to_taxid (org_key → taxid).
        Accessions already in _accdb are skipped.
        """
        import xml.etree.ElementTree as ET
        import urllib.parse as _ulp

        to_fetch = [a for a in accessions if a not in self._accdb]
        if not to_fetch:
            return

        api_key = self.cfg.get("api_key", "").strip()
        base    = self._NCBI_BASE

        for start in range(0, len(to_fetch), 500):
            if self._stop:
                return
            chunk  = to_fetch[start:start + 500]
            params: dict = {"db": "nucleotide", "id": ",".join(chunk),
                            "rettype": "gbc", "retmode": "xml"}
            if api_key:
                params["api_key"] = api_key
            xml_text = self._http_post(f"{base}efetch.fcgi",
                                       _ulp.urlencode(params), timeout=120)
            if not xml_text or self._stop:
                continue
            try:
                root = ET.fromstring(xml_text)
            except Exception:
                continue
            with self._cache_lock:
                for seq in root.iter("INSDSeq"):
                    accver   = (seq.findtext("INSDSeq_accession-version") or "").strip()
                    organism = (seq.findtext("INSDSeq_organism") or "").strip()
                    if not accver or not organism:
                        continue
                    org_key = (organism.replace(" ", "_")
                                       .replace("&apos;", "'")
                                       .replace("&amp;", "&"))
                    self._accdb[accver] = org_key
                    # Extract taxid from the source feature db_xref qualifier
                    for feat in seq.iter("INSDFeature"):
                        if feat.findtext("INSDFeature_key", "") == "source":
                            for qual in feat.iter("INSDQualifier"):
                                if qual.findtext("INSDQualifier_name", "") == "db_xref":
                                    val = qual.findtext("INSDQualifier_value", "")
                                    if val.startswith("taxon:"):
                                        self._org_to_taxid[org_key] = val.split(":")[1].strip()
                                        break
                            break

    def _fetch_taxonomy_batch(self, organisms: List[str]) -> None:
        """Fetch taxonomy for all organisms using taxids when available.

        Fast path: organisms in _org_to_taxid → single batch efetch on taxonomy DB.
        Fallback: organisms without a known taxid use the existing esearch-based
        _fetch_taxonomy() (rare — only for accessions loaded from .dbx cache files
        created by a previous run before this batch optimisation was added).
        Populates _taxadb (org_key → tab-separated ranks).
        """
        import urllib.parse as _ulp

        to_fetch = [o for o in organisms if o not in self._taxadb]
        if not to_fetch:
            return

        api_key = self.cfg.get("api_key", "").strip()
        base    = self._NCBI_BASE

        with_taxid    = [(o, self._org_to_taxid[o]) for o in to_fetch
                         if o in self._org_to_taxid]
        without_taxid = [o for o in to_fetch if o not in self._org_to_taxid]

        # ── Fast path: batch efetch by taxid ─────────────────────────────
        unique_taxids = list(dict.fromkeys(tid for _, tid in with_taxid))
        taxid_results: Dict[str, str] = {}
        # Taxids whose efetch chunk did NOT return a usable response (network
        # failure, stop, etc.). A negative for these is transient, not confirmed.
        failed_taxids: set = set()
        for start in range(0, len(unique_taxids), 500):
            if self._stop:
                # Treat the rest as transient so they are retried, never persisted.
                failed_taxids.update(unique_taxids[start:])
                break
            chunk  = unique_taxids[start:start + 500]
            params: dict = {"db": "taxonomy", "id": ",".join(chunk), "retmode": "xml"}
            if api_key:
                params["api_key"] = api_key
            xml_text = self._http_post(f"{base}efetch.fcgi",
                                       _ulp.urlencode(params), timeout=120)
            if xml_text and not self._stop:
                taxid_results.update(self._parse_taxonomy_xml_batch(xml_text))
            else:
                failed_taxids.update(chunk)

        with self._cache_lock:
            for org, taxid in with_taxid:
                if taxid in taxid_results:
                    # Confirmed answer (lineage found) from a successful response.
                    self._taxadb[org] = taxid_results[taxid]
                    self._tax_unconfirmed.discard(org)
                elif taxid in failed_taxids:
                    # Transient failure: keep a negative in memory so the in-run
                    # retry loop still finds it, but mark it unconfirmed so it is
                    # excluded from persistence to taxadb.dbx.
                    self._taxadb[org] = "Not_found_in_Taxonomy"
                    self._tax_unconfirmed.add(org)
                else:
                    # Chunk succeeded but this taxid was absent from the response =
                    # genuine "no lineage" negative → safe to cache and persist.
                    self._taxadb[org] = "Not_found_in_Taxonomy"
                    self._tax_unconfirmed.discard(org)

        # ── Fallback: esearch by scientific name (legacy cache hits) ──────
        for org in without_taxid:
            if self._stop:
                return
            self._fetch_taxonomy(org)

    def _parse_taxonomy_xml_batch(self, xml_text: str) -> Dict[str, str]:
        """Parse a multi-taxon taxonomy XML; returns {taxid_str: tab-separated ranks}."""
        import xml.etree.ElementTree as ET
        try:
            root = ET.fromstring(xml_text)
        except Exception:
            return {}
        ranks_wanted = {"kingdom", "class", "order", "family", "genus"}
        results: Dict[str, str] = {}
        for taxon_el in root.findall("Taxon"):   # direct children = one per requested taxid
            taxid = (taxon_el.findtext("TaxId") or "").strip()
            if not taxid:
                continue
            found: Dict[str, str] = {}
            for child in taxon_el.iter("Taxon"):  # includes LineageEx descendants
                rank = (child.findtext("Rank") or "").strip().lower()
                name = (child.findtext("ScientificName") or "").strip()
                if rank in ranks_wanted and rank not in found:
                    found[rank] = name
            lineage = "\t".join(
                found.get(r, "-") for r in ["kingdom", "class", "order", "family", "genus"]
            )
            results[taxid] = lineage
            # NCBI returns a merged taxon under its CURRENT TaxId; a requested taxid
            # that was merged is listed under <AkaTaxIds>. Map those too so the
            # caller's lookup by the requested id resolves instead of being wrongly
            # cached/persisted as "Not_found_in_Taxonomy".
            aka = taxon_el.find("AkaTaxIds")
            if aka is not None:
                for alt_el in aka.findall("TaxId"):
                    alt = (alt_el.text or "").strip()
                    if alt:
                        results[alt] = lineage
        return results

    # ── Cache I/O ────────────────────────────────────────────────────────

    def _load_cache(self, path):
        cache = {}
        if os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as fh:
                    for line in fh:
                        parts = line.rstrip("\n").split("\t", 1)
                        if len(parts) == 2:
                            cache[parts[0]] = parts[1]
            except Exception:
                pass
        return cache

    def _append_cache(self, new_entries: dict, path: str):
        """Append only new key-value pairs to the cache file (no full rewrite)."""
        if not new_entries:
            return
        try:
            with open(path, "a", encoding="utf-8") as fh:
                for k in sorted(new_entries):
                    fh.write(f"{k}\t{new_entries[k]}\n")
        except Exception:
            pass

    # ── TSV → XLSX conversion ─────────────────────────────────────────────

    def _tsv_to_xlsx(self, tsv_path: str) -> str:
        """Convert *tsv_path* to a formatted xlsx. Returns xlsx path or '' on failure."""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
        except ImportError as e:
            self.statusUpdated.emit("result", f"XLSX skip  │ openpyxl not available: {e}")
            return ""
        try:
            with open(tsv_path, "r", encoding="utf-8") as fh:
                lines = [l for l in fh.read().splitlines() if l.strip()]
        except Exception as e:
            self.statusUpdated.emit("result", f"XLSX skip  │ could not read TSV: {e}")
            return ""
        if not lines:
            return ""

        # The TSV already carries Hit_rank (per-sample 1..N, best hit = 1) as its
        # first column, so the layout is read as-is.
        headers = lines[0].split("\t")
        n_cols  = len(headers)

        # openpyxl 3.x requires 8-char ARGB hex strings (alpha + RGB).
        # Zone 1 → cols 0-1   : Hit_rank, Query_name
        # Zone 2 → cols 2-12  : BLAST metric columns (accession … bit-score)
        # Zone 3 → cols 13+   : Taxonomy columns (only when fetch_taxonomy=True)
        H1 = PatternFill(patternType="solid", fgColor="FF1A365D")   # header navy
        H2 = PatternFill(patternType="solid", fgColor="FF0D5E6E")   # header teal
        H3 = PatternFill(patternType="solid", fgColor="FF7C3200")   # header burnt-orange
        D1 = PatternFill(patternType="solid", fgColor="FFE8F1FB")   # data light-blue
        D2 = PatternFill(patternType="solid", fgColor="FFE8F5F6")   # data light-teal
        D3 = PatternFill(patternType="solid", fgColor="FFFEF3E8")   # data light-orange

        white_bold  = Font(color="FFFFFFFF", bold=True, size=10)
        normal_font = Font(size=10)
        bold_font   = Font(size=10, bold=True)   # best hit (Hit_rank == 1) per sample
        hdr_align   = Alignment(horizontal="center", vertical="center")
        dat_align   = Alignment(vertical="center", wrap_text=False)
        thin        = Side(style="thin", color="FFCCCCCC")
        medium      = Side(style="medium", color="FF9AA0A6")   # sample-block divider
        border      = Border(left=thin, right=thin, top=thin, bottom=thin)
        # Same as `border` but with a heavier top edge — marks the first row of each
        # new sample block so the hit groups are visually separated.
        border_group_top = Border(left=thin, right=thin, top=medium, bottom=thin)

        # Hit_rank and the BLAST metric columns (P_identity … Bit_score) are
        # numeric. Writing them as strings makes Excel flag every cell with
        # "Number stored as text", whose background error-checker re-scans the
        # sheet on every sort/filter/scroll → high CPU. Convert these to
        # int/float so openpyxl writes native numeric cells; the rest stay text.
        # Detect them by header name so it is robust to column shifts.
        _NUMERIC_NAMES = frozenset({
            "Hit_rank", "P_identity", "Alignment_length", "Num_mismatches",
            "Gap_opens", "Query_start", "Query_end", "Subject_start",
            "Subject_end", "Evalue", "Bit_score",
        })
        _numeric_idx = frozenset(
            i for i, h in enumerate(headers) if h in _NUMERIC_NAMES)

        def _num(v):
            if v == "":
                return v
            try:
                return int(v)
            except ValueError:
                pass
            try:
                return float(v)   # handles decimals and e-notation (Evalue)
            except ValueError:
                return v          # leave genuinely non-numeric text as-is

        def _hfill(ci):
            return H1 if ci <= 1 else (H2 if ci <= 12 else H3)

        def _dfill(ci, tinted):
            if not tinted:
                return None
            return D1 if ci <= 1 else (D2 if ci <= 12 else D3)

        wb = Workbook()
        ws = wb.active
        ws.title = "BLAST Results"

        # Header row
        ws.append(headers)
        for ci in range(n_cols):
            cell = ws.cell(row=1, column=ci + 1)
            cell.fill      = _hfill(ci)
            cell.font      = white_bold
            cell.alignment = hdr_align
            cell.border    = border
        ws.row_dimensions[1].height = 22

        # Data rows — shaded in BLOCKS by sample (Query_name), not per single row,
        # so each sample's hits share one tint and the next sample flips tint. A
        # heavier top border marks the first row of each new block. Rows already
        # arrive grouped by Query_name (see _rank_rows), so a simple value-change
        # check delimits the blocks. Column indices are looked up by name to stay
        # robust to layout shifts (taxonomy columns present or not).
        try:
            _query_idx = headers.index("Query_name")
        except ValueError:
            _query_idx = 1
        try:
            _rank_idx = headers.index("Hit_rank")
        except ValueError:
            _rank_idx = 0

        prev_query   = None
        block_tinted = True   # first block tinted (matches former row-2 behaviour)
        for rn, line in enumerate(lines[1:], start=2):
            raw_vals = line.split("\t")
            while len(raw_vals) < n_cols:
                raw_vals.append("")
            raw_vals = raw_vals[:n_cols]

            query     = raw_vals[_query_idx] if _query_idx < n_cols else ""
            new_block = (prev_query is not None and query != prev_query)
            if new_block:
                block_tinted = not block_tinted
            prev_query = query

            rank_raw = raw_vals[_rank_idx] if _rank_idx < n_cols else ""
            is_best  = (str(rank_raw).strip() == "1")   # best hit of this sample

            vals = [
                _num(v) if ci in _numeric_idx else v
                for ci, v in enumerate(raw_vals)
            ]
            ws.append(vals)
            row_border = border_group_top if new_block else border
            row_font   = bold_font if is_best else normal_font
            for ci in range(n_cols):
                cell = ws.cell(row=rn, column=ci + 1)
                fill = _dfill(ci, block_tinted)
                if fill:
                    cell.fill = fill
                cell.font      = row_font
                cell.alignment = dat_align
                cell.border    = row_border

        # Freeze header, enable auto-filter
        ws.freeze_panes    = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Auto-fit column widths (capped at 55 chars)
        for ci, col_cells in enumerate(ws.columns):
            width = max((len(str(c.value or "")) for c in col_cells), default=8)
            ws.column_dimensions[get_column_letter(ci + 1)].width = min(width + 2, 55)

        xlsx_path = tsv_path.rsplit(".", 1)[0] + ".xlsx"
        try:
            wb.save(xlsx_path)
        except Exception as e:
            self.statusUpdated.emit("result", f"XLSX error │ {e}")
            return ""
        return xlsx_path

    # ── Main run ─────────────────────────────────────────────────────────

    def run(self):
        try:
            self._run_blast()
        except Exception as e:
            import traceback
            self.taskError.emit(f"{e}\n{traceback.format_exc()}")

    def _run_blast(self):
        cfg       = self.cfg
        nhits     = cfg["nhits"]
        nseq      = cfg["nseq"]
        run_start = datetime.datetime.now()
        mydate    = run_start.strftime("%Y%m%d-%H%M%S")

        # ── Output directory (passed from MainWindow dialog) ──
        output_dir = cfg["outdir"]
        os.makedirs(output_dir, exist_ok=True)

        # ── Cache files stored alongside the output folder ──
        taxadb_path = os.path.join(output_dir, "taxadb.dbx")
        accdb_path  = os.path.join(output_dir, "accdb.dbx")
        fetch_tax = cfg.get("fetch_taxonomy", True)
        if fetch_tax:
            # Pre-load any .dbx files from sibling blast result folders
            parent = os.path.dirname(output_dir)
            for root_dir, _dirs, fnames in os.walk(parent):
                if "taxadb.dbx" in fnames:
                    self._taxadb.update(
                        self._load_cache(os.path.join(root_dir, "taxadb.dbx"))
                    )
                if "accdb.dbx" in fnames:
                    self._accdb.update(
                        self._load_cache(os.path.join(root_dir, "accdb.dbx"))
                    )
            # Snapshot keys already on disk — only new ones will be appended
            self._saved_tax_keys = set(self._taxadb.keys())
            self._saved_acc_keys = set(self._accdb.keys())

        # ── Merge & normalize FASTA ──
        raw   = self._merge_fasta_files(self.files)
        fasta = self._to_single_line_fasta(raw)
        seq_count = fasta.count("\n>") + (1 if fasta.startswith(">") else 0)

        if seq_count == 0:
            self.taskError.emit(
                "No FASTA sequences found in the provided files."
            )
            return

        # ── Split into batches ──
        batches, batch_pairs_list = self._split_batches(fasta, nseq)
        n_batches = len(batches)
        completed_batches: set = set()   # indices of batches fully written to TSV

        # ── Summary line (fixed slot "info") ──
        self.statusUpdated.emit(
            "info",
            f"Sequences: {seq_count}  │  Batches: {n_batches}"
            f"  │  Hits/seq: {nhits}  │  DB: {cfg['database']}"
        )

        _blast_cols = (
            "Query_name\tSubject_accession.ver\tP_identity\tAlignment_length\t"
            "Num_mismatches\tGap_opens\tQuery_start\tQuery_end\t"
            "Subject_start\tSubject_end\tEvalue\tBit_score"
        )
        headings = (
            _blast_cols + "\tSubject_Kingdom\tSubject_Class\tSubject_Order\t"
            "Subject_Family\tSubject_Genus\tSubject_organism"
            if fetch_tax else _blast_cols
        )
        # Per-sample hit rank (1 = best hit) as the first column, so results can
        # be filtered by rank (e.g. Hit_rank == 1 keeps only each sample's top hit).
        headings = "Hit_rank\t" + headings

        # ── Open TSV for incremental writing ──
        tsv_path = os.path.join(output_dir, f"blast-{mydate}.tsv")
        for _attempt in range(10):
            try:
                with open(tsv_path, "w", encoding="utf-8") as tsv_fh:
                    tsv_fh.write(headings + "\n")
                break
            except PermissionError:
                self.statusUpdated.emit(
                    "result",
                    f"⚠ Output file locked — close it to continue… ({_attempt + 1}/10)"
                )
                self._interruptible_sleep(0.5)
        else:
            # Sin la cabecera no tiene sentido lanzar las consultas BLAST (red):
            # se produciría un TSV sin encabezado. Abortar limpiamente.
            self.taskError.emit(
                f"Could not write output file (locked/permission denied):\n{tsv_path}"
            )
            return

        total_hits_done: int = 0
        total_expected = n_batches * self._BATCH_UNITS

        for batch_idx, batch_fasta in enumerate(batches):
            if self._stop:
                break

            batch_seq   = batch_fasta.count(">")
            batch_label = f"Batch {batch_idx+1}/{n_batches}"
            batch_base  = batch_idx * self._BATCH_UNITS

            # ── BLAST ──
            self.statusUpdated.emit(
                "blast",
                f"BLAST       │ [{batch_label}] Submitting {batch_seq} sequences…"
            )
            rid, rtoe = self._blast_submit(batch_fasta)
            if not rid:
                self.statusUpdated.emit(
                    "blast",
                    f"BLAST       │ [{batch_label}] Submission failed — skipping."
                )
                continue

            self.statusUpdated.emit(
                "blast",
                f"BLAST       │ [{batch_label}] RID={rid}  waiting {rtoe}s…"
            )
            for tick in range(rtoe):
                if self._stop:
                    break
                time.sleep(1)
                self.progressUpdated.emit(
                    batch_base + int(400 * (tick + 1) / max(rtoe, 1)),
                    total_expected
                )
            if self._stop:
                break

            self.progressUpdated.emit(batch_base + 400, total_expected)
            if not self._blast_poll(rid, batch_label):
                continue

            blast_rows = self._blast_get_tabular(rid)
            self.statusUpdated.emit(
                "blast",
                f"BLAST       │ [{batch_label}] {len(blast_rows)} hits retrieved  ✓"
            )

            if fetch_tax:
                # ── Fetch organisms (batch: one POST for all accessions) ──
                accessions = [
                    (row.split("\t")[1] if "\t" in row else "") for row in blast_rows
                ]
                unique_accs   = list(dict.fromkeys(a for a in accessions if a))
                n_unique_accs = len(unique_accs)
                self.statusUpdated.emit(
                    "organism",
                    f"Organism ID │ [{batch_label}] Fetching {n_unique_accs} accessions…"
                )
                self._fetch_organisms_batch(unique_accs)
                if self._stop:
                    break
                n_org_found = sum(1 for a in unique_accs if a in self._accdb)
                self.statusUpdated.emit(
                    "organism",
                    f"Organism ID │ [{batch_label}] {n_org_found}/{n_unique_accs} resolved  ✓"
                )
                self.progressUpdated.emit(batch_base + 600, total_expected)

                organisms: List[str] = [
                    (self._accdb.get(acc, "") if acc else "") for acc in accessions
                ]

                # ── Fetch taxonomy (batch: one POST per 500 taxids) ───────
                unique_orgs   = list(dict.fromkeys(o for o in organisms if o))
                n_unique_orgs = len(unique_orgs)
                self.statusUpdated.emit(
                    "taxonomy",
                    f"Taxonomy    │ [{batch_label}] Fetching {n_unique_orgs} organisms…"
                )
                self._fetch_taxonomy_batch(unique_orgs)
                if self._stop:
                    break

                # ── Retry not-found organisms (batch) ─────────────────────
                for attempt in range(1, self._TAX_RETRIES + 1):
                    if self._stop:
                        break
                    retry_orgs = [
                        org for org in unique_orgs
                        if self._taxadb.get(org) == "Not_found_in_Taxonomy"
                    ]
                    if not retry_orgs:
                        break
                    n_retry = len(retry_orgs)
                    with self._cache_lock:
                        for org in retry_orgs:
                            del self._taxadb[org]
                    self.statusUpdated.emit(
                        "taxonomy",
                        f"Taxonomy    │ [{batch_label}] Retry {attempt}/{self._TAX_RETRIES}: {n_retry} not-found…"
                    )
                    self._fetch_taxonomy_batch(retry_orgs)

                n_tax_found = sum(
                    1 for o in unique_orgs
                    if self._taxadb.get(o, "Not_found_in_Taxonomy") != "Not_found_in_Taxonomy"
                )
                self.statusUpdated.emit(
                    "taxonomy",
                    f"Taxonomy    │ [{batch_label}] {n_tax_found}/{n_unique_orgs} resolved  ✓"
                )
                self.progressUpdated.emit(batch_base + 800, total_expected)

                taxonomies: List[str] = [
                    (self._taxadb.get(o, "Not_found_in_Taxonomy") if o else "Not_found_in_Taxonomy")
                    for o in organisms
                ]

                batch_rows_out = self._rank_rows([
                    f"{row}\t{tax}\t{org}"
                    for row, tax, org in zip(blast_rows, taxonomies, organisms)
                ])

                # ── Persist only new entries after every batch ──
                # Exclude _tax_unconfirmed: negatives from a failed efetch are kept
                # in memory for the in-run retry but must not poison taxadb.dbx.
                with self._cache_lock:
                    new_tax = {k: self._taxadb[k] for k in self._taxadb
                               if k not in self._saved_tax_keys
                               and k not in self._tax_unconfirmed}
                    new_acc = {k: self._accdb[k]  for k in self._accdb  if k not in self._saved_acc_keys}
                    self._append_cache(new_tax, taxadb_path)
                    self._append_cache(new_acc,  accdb_path)
                    self._saved_tax_keys.update(new_tax.keys())
                    self._saved_acc_keys.update(new_acc.keys())

            else:
                batch_rows_out = self._rank_rows(blast_rows)

            # ── Append batch rows to disk immediately ──
            row_phase_start = 800 if fetch_tax else 400
            row_phase_range = 200 if fetch_tax else 600
            n_batch_rows = max(len(batch_rows_out), 1)
            _batch_written = False
            for _attempt in range(20):   # retry up to 10 s if TSV is open in Excel
                try:
                    with open(tsv_path, "a", encoding="utf-8") as tsv_fh:
                        for row_idx, r in enumerate(batch_rows_out, 1):
                            tsv_fh.write(r + "\n")
                            self.progressUpdated.emit(
                                batch_base + row_phase_start + int(row_phase_range * row_idx / n_batch_rows),
                                total_expected
                            )
                    _batch_written = True
                    break  # write succeeded
                except PermissionError:
                    self.statusUpdated.emit(
                        "result",
                        f"⚠ TSV file is open — close it and the run will resume… ({_attempt + 1}/20)"
                    )
                    self._interruptible_sleep(0.5)

            # Solo marcar el batch como completado si sus filas llegaron al disco;
            # de lo contrario sus secuencias deben aparecer en el FASTA de "missing".
            # Contar los hits una sola vez, tras la escritura exitosa, para que un
            # reintento (que reescribe el lote completo) no infle el total.
            if _batch_written:
                completed_batches.add(batch_idx)
                total_hits_done += len(batch_rows_out)

        # ── Build missing-sequences FASTA (unprocessed or failed batches) ──
        missing_pairs = []
        for bi, bp in enumerate(batch_pairs_list):
            if bi not in completed_batches:
                missing_pairs.extend(bp)

        miss_msg = ""
        if missing_pairs:
            miss_path = os.path.join(output_dir, f"missing_seqs_{mydate}.fa")
            try:
                with open(miss_path, "w", encoding="utf-8") as fh:
                    for h, s in missing_pairs:
                        fh.write(h + "\n" + s + "\n")
                miss_msg = (
                    f"{len(missing_pairs)} unprocessed seqs → "
                    f"{os.path.basename(miss_path)}"
                )
            except Exception as exc:
                miss_msg = f"Could not write missing FASTA: {exc}"

        # ── Convert TSV → XLSX ──
        xlsx_path = ""
        if not self._stop and total_hits_done > 0:
            xlsx_path = self._tsv_to_xlsx(tsv_path)

        if self._stop:
            result_msg = f"Stopped     │ {miss_msg}" if miss_msg else "Stopped by user."
        else:
            out_name = os.path.basename(xlsx_path if xlsx_path else tsv_path)
            if miss_msg:
                result_msg = f"Done  ✓     │ {total_hits_done} hits written  │ {miss_msg}"
            else:
                result_msg = f"Done  ✓     │ {total_hits_done} hits written → {out_name}"
        self.statusUpdated.emit("result", result_msg)

        # ── Write run log ──────────────────────────────────────────────────
        elapsed   = datetime.datetime.now() - run_start
        total_sec = int(elapsed.total_seconds())
        h, rem    = divmod(total_sec, 3600)
        m, s      = divmod(rem, 60)
        elapsed_str = f"{h}h {m:02d}m {s:02d}s" if h else f"{m:02d}m {s:02d}s"

        api_key = cfg.get("api_key", "")
        if len(api_key) > 8:
            api_masked = api_key[:4] + "*" * (len(api_key) - 8) + api_key[-4:]
        elif api_key:
            api_masked = "****"
        else:
            api_masked = "(not set)"

        status_str = "Stopped" if self._stop else "Completed"
        batches_ok = len(completed_batches)

        log_lines = [
            "BLAST Run Log",
            "=" * 60,
            f"Date/Time  : {run_start.strftime('%Y-%m-%d %H:%M:%S')}",
            f"Status     : {status_str}",
            f"Total time : {elapsed_str}",
            "",
            "Input files:",
        ]
        for f in self.files:
            log_lines.append(f"  {os.path.abspath(f)}")
        log_lines += [
            "",
            "Parameters:",
            f"  Database          : {cfg.get('database', '')}",
            f"  Program           : {cfg.get('program', '')}",
            f"  Hits per sequence : {nhits}",
            f"  Sequences / batch : {nseq}",
            f"  Fetch taxonomy    : {'Yes' if fetch_tax else 'No'}",
            f"  NCBI API key      : {api_masked}",
            "",
            "Results:",
            f"  Sequences found   : {seq_count}",
            f"  Batches           : {batches_ok}/{n_batches} completed",
            f"  Hits written      : {total_hits_done}",
            f"  Output folder     : {output_dir}",
            f"  TSV file          : {os.path.basename(tsv_path)}",
            f"  XLSX file         : {os.path.basename(xlsx_path) if xlsx_path else 'N/A'}",
        ]
        if miss_msg:
            log_lines.append(f"  Missing seqs      : {miss_msg}")
        log_lines += [
            "",
            "NOTE: Do not delete the .dbx cache files (accdb.dbx, taxadb.dbx).",
            "      They store organism and taxonomy lookups already performed and",
            "      will significantly speed up future BLAST runs on the same or",
            "      overlapping accession numbers.",
            "",
        ]

        os.makedirs(output_dir, exist_ok=True)
        log_path = os.path.join(output_dir, f"blast_run_log_{mydate}.txt")
        try:
            with open(log_path, "w", encoding="utf-8") as lf:
                lf.write("\n".join(log_lines))
        except Exception as exc:
            self.statusUpdated.emit("result", f"{result_msg}  │  Log error: {exc}")

        self.statusUpdated.emit(
            "info",
            f"Session stats  │  {self._req_count} requests · "
            f"{self._rate_limit_count} rate limit(s) (429) · "
            f"{self._server_err_count} server error(s) (5xx)"
        )
        self.taskFinished.emit(output_dir)



# ═══════════════════════════════════════════════════════════════════════════
# FASTA COMPARE
# ═══════════════════════════════════════════════════════════════════════════

class _CompareResultsWindow(QtWidgets.QDialog):
    """Separate window showing the table of comparison results."""

    _BG = {
        "Identical":          "#EAF3DE",
        "Compatible (IUPAC)": "#E6F1FB",
        "Different":          "#FCEBEB",
        "Only in reference":  "#FAEEDA",
        "No reference":       "#F5F5F3",
        "Unique":             "#EBEBEB",
    }
    _PILL_BG = {
        "Identical":          ("#3B6D11", "#EAF3DE"),
        "Compatible (IUPAC)": ("#185FA5", "#E6F1FB"),
        "Different":          ("#A32D2D", "#FCEBEB"),
        "Only in reference":  ("#854F0B", "#FAEEDA"),
        "No reference":       ("#5A5A5A", "#EBEBEB"),
        "Unique":             ("#4A4A4A", "#DCDCDC"),
    }

    def __init__(self, rows, headers, outdir, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Comparison results — ONTbarcoder")
        self.resize(1200, 680)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setStyleSheet(f"background-color: {WHITE}; font-family: 'Segoe UI', Arial, sans-serif;")

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        topbar = QtWidgets.QWidget()
        topbar.setStyleSheet("background-color: #1A1A2E;")
        topbar.setFixedHeight(62)
        tb_layout = QtWidgets.QHBoxLayout(topbar)
        tb_layout.setContentsMargins(16, 6, 16, 6)
        tb_layout.setSpacing(12)

        n_identical = sum(1 for r in rows if r.get("State") == "Identical")
        n_compat    = sum(1 for r in rows if r.get("State") == "Compatible (IUPAC)")
        n_diff      = sum(1 for r in rows if r.get("State") == "Different")
        n_only      = sum(1 for r in rows if r.get("State", "").startswith("Unique"))
        n_only_ref  = sum(1 for r in rows if r.get("State") == "Only in reference")
        n_no_ref    = sum(1 for r in rows if r.get("State") == "No reference")

        def _stat_pill(label, count, bg, fg):
            w = QtWidgets.QLabel(f"{label}: <b>{count}</b>")
            w.setTextFormat(QtCore.Qt.RichText)
            w.setStyleSheet(
                f"background-color:{bg}; color:{fg}; border-radius:8px;"
                f" padding:3px 10px; font-size:15px;"
            )
            return w

        title_lbl = QtWidgets.QLabel(
            f"<b style='color:white;font-size:18px;'>Comparison</b>"
            f"<span style='color:#9AAFCC;font-size:15px;'>  —  {len(rows)} IDs</span>")
        title_lbl.setTextFormat(QtCore.Qt.RichText)
        tb_layout.addWidget(title_lbl)
        tb_layout.addSpacing(10)

        pills = [
            ("Identical",        n_identical, "#2D6A0A", "#C5EAAB"),
            ("Compatible IUPAC", n_compat,    "#0D4D8A", "#A8CCF0"),
            ("Different",        n_diff,      "#8B1A1A", "#F5AAAA"),
            ("Unique",           n_only,      "#444444", "#CCCCCC"),
        ]
        for lbl, cnt, fg, bg in pills:
            if cnt:
                tb_layout.addWidget(_stat_pill(lbl, cnt, bg, fg))
        if n_only_ref:
            tb_layout.addWidget(_stat_pill("Only ref", n_only_ref, "#6B3D0A", "#F5D9A8"))
        if n_no_ref:
            tb_layout.addWidget(_stat_pill("No ref", n_no_ref, "#3A3A3A", "#DDDDDD"))
        tb_layout.addStretch()
        layout.addWidget(topbar)

        if outdir:
            path_bar = QtWidgets.QWidget()
            path_bar.setStyleSheet(f"background-color: #F0EEE8; border-bottom: 1px solid {GRAY_LINE};")
            path_bar.setFixedHeight(30)
            pb_layout = QtWidgets.QHBoxLayout(path_bar)
            pb_layout.setContentsMargins(20, 0, 16, 0)
            path_lbl = QtWidgets.QLabel(
                f"<span style='color:{TEXT_HINT};'>📁 Output:</span>"
                f" <span style='color:{TEXT_SEC};'>{outdir}</span>")
            path_lbl.setTextFormat(QtCore.Qt.RichText)
            path_lbl.setStyleSheet("font-size: 11px;")
            pb_layout.addWidget(path_lbl, 1)
            layout.addWidget(path_bar)

        table = QtWidgets.QTableWidget(len(rows), len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        table.setShowGrid(False)
        table.setAlternatingRowColors(False)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setHighlightSections(False)
        table.setFocusPolicy(QtCore.Qt.NoFocus)
        table.setStyleSheet(f"""
            QTableWidget {{
                background-color: {WHITE};
                border: none;
                outline: none;
                font-size: 13px;
                gridline-color: transparent;
            }}
            QTableWidget::item {{
                padding: 5px 10px;
                border-bottom: 1px solid {GRAY_LINE};
            }}
            QTableWidget::item:selected {{
                background-color: #CDE0F5;
                color: #1A1A2E;
            }}
            QHeaderView::section {{
                background-color: #1A1A2E;
                color: white;
                font-weight: 600;
                font-size: 12px;
                padding: 7px 10px;
                border: none;
                border-right: 1px solid #2E3A50;
            }}
            QHeaderView::section:last {{ border-right: none; }}
            QScrollBar:vertical {{ width: 8px; background: {GRAY_BG}; }}
            QScrollBar::handle:vertical {{
                background: {GRAY_LINE}; border-radius: 4px; min-height: 20px;
            }}
            QScrollBar:horizontal {{ height: 8px; background: {GRAY_BG}; }}
            QScrollBar::handle:horizontal {{
                background: {GRAY_LINE}; border-radius: 4px; min-width: 20px;
            }}
        """)
        table.horizontalHeader().setMinimumSectionSize(60)

        color_map = {
            "Identical":          QtGui.QColor("#EAF3DE"),
            "Compatible (IUPAC)": QtGui.QColor("#E6F1FB"),
            "Different":          QtGui.QColor("#FCEBEB"),
            "Only in reference":  QtGui.QColor("#FAEEDA"),
            "No reference":       QtGui.QColor("#F0F0EE"),
        }
        estado_col = headers.index("State") if "State" in headers else -1

        for i, row in enumerate(rows):
            table.setRowHeight(i, 32)
            estado = row.get("State", "")
            bg = color_map.get(estado)
            if not bg and estado.startswith("Unique"):
                bg = QtGui.QColor("#EBEBEB")
            for j, h in enumerate(headers):
                val = str(row.get(h, ""))
                item = QtWidgets.QTableWidgetItem(val)
                item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                if bg:
                    item.setBackground(QtGui.QBrush(bg))
                if h == "ID":
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                table.setItem(i, j, item)
            if estado_col >= 0:
                # "Unique in <file>" maps to the generic "Unique" pill; every
                # other state must match its key exactly.
                pill_key = "Unique" if estado.startswith("Unique") else (
                    estado if estado in self._PILL_BG else None)
                if pill_key:
                    fg_col, bg_col = self._PILL_BG[pill_key]
                    pill = QtWidgets.QLabel(f"  {estado}  ")
                    pill.setAlignment(QtCore.Qt.AlignCenter)
                    pill.setStyleSheet(
                        f"background-color:{bg_col}; color:{fg_col};"
                        f" border-radius:9px; font-size:11px; font-weight:600;"
                        f" padding:1px 6px;"
                    )
                    if bg:
                        container = QtWidgets.QWidget()
                        container.setStyleSheet(f"background-color:{bg.name()};")
                        cl = QtWidgets.QHBoxLayout(container)
                        cl.setContentsMargins(4, 2, 4, 2)
                        cl.addWidget(pill)
                        table.setCellWidget(i, estado_col, container)

        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table)

        footer = QtWidgets.QWidget()
        footer.setStyleSheet(f"background-color:{GRAY_BG}; border-top:1px solid {GRAY_LINE};")
        footer.setFixedHeight(52)
        fl = QtWidgets.QHBoxLayout(footer)
        fl.setContentsMargins(20, 8, 20, 8)
        fl.addStretch()
        if outdir:
            open_btn = QtWidgets.QPushButton("Open folder  📂")
            open_btn.setFixedHeight(36)
            open_btn.setStyleSheet(
                f"QPushButton {{ background:{BLUE_LIGHT}; color:{BLUE}; border:1px solid #B8D4F0;"
                f" border-radius:7px; padding:4px 14px; font-size:13px; }}"
                f"QPushButton:hover {{ background:{BLUE}; color:white; }}"
            )
            open_btn.clicked.connect(lambda: QtGui.QDesktopServices.openUrl(
                QtCore.QUrl.fromLocalFile(outdir)))
            fl.addWidget(open_btn)
            fl.addSpacing(8)
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.setFixedHeight(36)
        close_btn.setStyleSheet(
            f"QPushButton {{ background:{GRAY_LINE}; color:{TEXT_SEC}; border:none;"
            f" border-radius:7px; padding:4px 18px; font-size:13px; }}"
            f"QPushButton:hover {{ background:#C8C6C0; color:{TEXT_PRI}; }}"
        )
        close_btn.clicked.connect(self.close)
        fl.addWidget(close_btn)
        layout.addWidget(footer)


class _CompareWorker(QtCore.QThread):
    """Compare N files with each other (all possible pairs)."""
    notifyProgress = QtCore.pyqtSignal(int)
    taskFinished   = QtCore.pyqtSignal(list, list)

    def __init__(self, file_list: List[str], outdir: str, extract_cfg=None,
                 parent=None):
        super().__init__(parent)
        self.file_list   = file_list
        self.outdir      = outdir
        self.extract_cfg = extract_cfg

    def run(self):
        file_list = self.file_list
        basenames = _make_unique_labels(file_list)

        seqs: Dict[str, Dict[str, Tuple]] = {}
        for fname, bn in zip(file_list, basenames):
            seqs[bn] = _parse_fasta_file(fname, self.extract_cfg)

        pairs   = list(itertools.combinations(basenames, 2))
        all_ids = sorted({sid for d in seqs.values() for sid in d})
        total   = len(all_ids)

        rows: List[dict] = []
        priority = {"different": 0, "compatible": 1, "identical": 2}

        for prog_i, sid in enumerate(all_ids):
            present_in  = [bn for bn in basenames if sid in seqs.get(bn, {})]
            absent_from = [bn for bn in basenames if sid not in seqs.get(bn, {})]
            row: dict = {"ID": sid}

            for bn in basenames:
                entry = seqs.get(bn, {}).get(sid)
                if entry:
                    row[f"{bn}_len"]  = entry[1]
                    row[f"{bn}_cov"]  = entry[2]
                    row[f"{bn}_ambs"] = entry[3]
                    row[f"{bn}_gaps"] = entry[4]
                else:
                    row[f"{bn}_len"] = row[f"{bn}_cov"] = \
                    row[f"{bn}_ambs"] = row[f"{bn}_gaps"] = ""

            if len(present_in) == 1:
                row["State"]      = f"Unique in {present_in[0]}"
                row["Best_run"]   = present_in[0]
                row["Diff_bases"] = ""
                row["Note"]       = "Not found in other runs"
                rows.append(row)
                self.notifyProgress.emit(int((prog_i + 1) / total * 100))
                continue

            pair_results: Dict[Tuple[str, str], Tuple[str, int, int, bool]] = {}
            for bn1, bn2 in pairs:
                if sid not in seqs.get(bn1, {}) or sid not in seqs.get(bn2, {}):
                    continue
                s1 = seqs[bn1][sid][0]
                s2 = seqs[bn2][sid][0]
                estado_par, d_amb, d_noamb, rc_used = _compare_sequences(s1, s2)
                pair_results[(bn1, bn2)] = (estado_par, d_amb, d_noamb, rc_used)

            lens = {bn: len(seqs[bn][sid][0]) for bn in present_in}
            unique_lens = sorted(set(lens.values()))
            if len(unique_lens) > 1:
                if len(present_in) == 2:
                    bns = list(lens.keys())
                    l0, l1_ = lens[bns[0]], lens[bns[1]]
                    delta = abs(l0 - l1_)
                    sym = "<" if l0 < l1_ else ">"
                    row["Length"] = f"{bns[0]} {sym} {bns[1]}: Δ{delta} bp"
                else:
                    row["Length"] = f"{unique_lens[0]}–{unique_lens[-1]} bp"
            else:
                row["Length"] = ""

            estados_pares = [v[0] for v in pair_results.values()]
            global_raw = min(estados_pares, key=lambda e: priority.get(e, 99)) \
                         if estados_pares else 'identical'

            estado_display = {
                'identical':  'Identical',
                'compatible': 'Compatible (IUPAC)',
                'different':  'Different',
            }.get(global_raw, global_raw)
            row["State"] = estado_display

            detail_parts = []
            max_dist = max_d_noamb = 0
            for (bn1, bn2), (ep, d_amb, d_noamb, rc_used) in pair_results.items():
                ep_es = {'identical': 'Identical', 'compatible': 'Compatible (IUPAC)',
                         'different': 'Different'}.get(ep, ep)
                nota_par = f"{bn1}↔{bn2}: {ep_es}"
                if ep == 'different' and d_amb:
                    nota_par += f" (d={d_amb})"
                    if d_amb > max_dist:
                        max_dist = d_amb
                elif ep == 'compatible' and d_noamb:
                    nota_par += f" (IUPAC_pos={d_noamb})"
                    if d_noamb > max_d_noamb:
                        max_d_noamb = d_noamb
                if rc_used:
                    nota_par += " [RC]"
                detail_parts.append(nota_par)

            row["Diff_bases"] = max_dist    if max_dist    else ""
            row["IUPAC_pos"]  = max_d_noamb if max_d_noamb else ""

            seqs_present: List[_SeqEntry] = [
                (bn, seqs[bn][sid][0], seqs[bn][sid][1],
                 seqs[bn][sid][2], seqs[bn][sid][3], seqs[bn][sid][4])
                for bn in present_in
            ]
            best = _best_barcode(seqs_present)
            if global_raw == 'different':
                if best:
                    row["Best_run"] = best[0]
                    detail_parts.append(
                        f"Best: {best[0]} (ambs={best[4]}, gaps={best[5]}, cov={best[3]})")
                else:
                    row["Best_run"] = "Tie"
                    detail_parts.append("Tie in quality")
            else:
                row["Best_run"] = ""

            if absent_from:
                detail_parts.append(f"Absent from: {', '.join(absent_from)}")

            row["Note"] = " | ".join(detail_parts)
            rows.append(row)
            self.notifyProgress.emit(int((prog_i + 1) / total * 100))

        has_iupac  = any(row.get("IUPAC_pos", "") for row in rows)
        has_length = any(row.get("Length", "")    for row in rows)
        headers = ["ID", "State", "Best_run", "Diff_bases"]
        if has_iupac:
            headers.append("IUPAC_pos")
        if has_length:
            headers.append("Length")
        for bn in basenames:
            headers += [f"{bn}_len", f"{bn}_cov", f"{bn}_ambs", f"{bn}_gaps"]
        headers.append("Note")

        _write_outputs(rows, headers, self.outdir,
                       all_bns=basenames, ref_bn=None, seqs_store=seqs)
        self.taskFinished.emit(rows, headers)


class _PairCompareWorker(QtCore.QThread):
    """Compares N files against a reference file chosen by the user."""
    notifyProgress = QtCore.pyqtSignal(int)
    taskFinished   = QtCore.pyqtSignal(list, list)

    def __init__(self, file_list: List[str], ref_path: str,
                 outdir: str, extract_cfg=None, parent=None):
        super().__init__(parent)
        self.file_list   = file_list
        self.ref_path    = ref_path
        self.outdir      = outdir
        self.extract_cfg = extract_cfg

    def run(self):
        file_list = self.file_list
        all_bns   = _make_unique_labels(file_list)

        ref_bn = next(
            (lbl for f, lbl in zip(file_list, all_bns)
             if os.path.abspath(f) == os.path.abspath(self.ref_path)),
            all_bns[0],
        )
        comp_bns = [bn for bn in all_bns if bn != ref_bn]

        seqs: Dict[str, Dict[str, Tuple]] = {}
        for fname, bn in zip(file_list, all_bns):
            seqs[bn] = _parse_fasta_file(fname, self.extract_cfg)

        ref_seqs = seqs.get(ref_bn, {})
        all_ids  = sorted(
            set(ref_seqs.keys()) |
            {sid for bn in comp_bns for sid in seqs.get(bn, {}).keys()}
        )
        total = len(all_ids)
        rows: List[dict] = []
        priority = {"different": 0, "compatible": 1, "identical": 2}

        for prog_i, sid in enumerate(all_ids):
            row: dict = {"ID": sid}

            if sid in ref_seqs:
                r = ref_seqs[sid]
                row[f"{ref_bn}_len"]  = r[1]
                row[f"{ref_bn}_cov"]  = r[2]
                row[f"{ref_bn}_ambs"] = r[3]
                row[f"{ref_bn}_gaps"] = r[4]
                ref_seq = r[0]
            else:
                row[f"{ref_bn}_len"] = row[f"{ref_bn}_cov"] = \
                row[f"{ref_bn}_ambs"] = row[f"{ref_bn}_gaps"] = ""
                ref_seq = None

            for bn in comp_bns:
                entry = seqs.get(bn, {}).get(sid)
                if entry:
                    row[f"{bn}_len"]  = entry[1]
                    row[f"{bn}_cov"]  = entry[2]
                    row[f"{bn}_ambs"] = entry[3]
                    row[f"{bn}_gaps"] = entry[4]
                else:
                    row[f"{bn}_len"] = row[f"{bn}_cov"] = \
                    row[f"{bn}_ambs"] = row[f"{bn}_gaps"] = ""
                row[f"{bn}_dif"] = ""

            if ref_seq is not None and all(
                sid not in seqs.get(bn, {}) for bn in comp_bns
            ):
                row["State"]      = "Only in reference"
                row["Best_run"]   = ref_bn
                row["Diff_bases"] = ""
                row["Note"]       = "Not found in any compared file"
                rows.append(row)
                self.notifyProgress.emit(int((prog_i + 1) / total * 100))
                continue

            if ref_seq is None:
                present_bns = [bn for bn in comp_bns if sid in seqs.get(bn, {})]
                row["State"]      = "No reference"
                row["Best_run"]   = ""
                row["Diff_bases"] = ""
                row["Note"]       = f"Only in: {', '.join(present_bns)}"
                rows.append(row)
                self.notifyProgress.emit(int((prog_i + 1) / total * 100))
                continue

            estados_por_bn:  Dict[str, str]  = {}
            dists_por_bn:    Dict[str, int]  = {}
            noamb_por_bn:    Dict[str, int]  = {}
            rc_por_bn:       Dict[str, bool] = {}
            diff_candidates: List[_SeqEntry] = []
            max_dist = max_d_noamb = 0

            for bn in comp_bns:
                if sid not in seqs.get(bn, {}):
                    estados_por_bn[bn] = "Absent"
                    dists_por_bn[bn]   = 0
                    noamb_por_bn[bn]   = 0
                    rc_por_bn[bn]      = False
                    row[f"{bn}_length"] = ""
                    continue

                comp_seq = seqs[bn][sid][0]
                estado_raw, d_amb, d_noamb, rc_used = _compare_sequences(ref_seq, comp_seq)

                estados_por_bn[bn] = estado_raw
                dists_por_bn[bn]   = d_amb
                noamb_por_bn[bn]   = d_noamb
                rc_por_bn[bn]      = rc_used

                row[f"{bn}_dif"] = d_amb if d_amb else ""
                if d_amb and d_amb > max_dist:
                    max_dist = d_amb
                if estado_raw == 'compatible' and d_noamb > max_d_noamb:
                    max_d_noamb = d_noamb

                delta = len(comp_seq) - len(ref_seq)
                if delta == 0:
                    row[f"{bn}_length"] = ""
                elif delta > 0:
                    row[f"{bn}_length"] = f"> {delta} bp"
                else:
                    row[f"{bn}_length"] = f"< {-delta} bp"

                if estado_raw == 'different':
                    e = seqs[bn][sid]
                    diff_candidates.append((bn, e[0], e[1], e[2], e[3], e[4]))

            length_parts = [
                f"{bn}: {row[f'{bn}_length']}"
                for bn in comp_bns
                if row.get(f"{bn}_length", "")
            ]
            row["Length"] = " | ".join(length_parts) if length_parts else ""

            valid = [e for e in estados_por_bn.values() if e != "Absent"]
            global_raw = min(valid, key=lambda e: priority.get(e, 99)) if valid else "Absent in all"

            estado_display = {
                'identical':  'Identical',
                'compatible': 'Compatible (IUPAC)',
                'different':  'Different',
            }.get(global_raw, global_raw)
            row["State"]      = estado_display
            row["Diff_bases"] = max_dist    if max_dist    else ""
            row["IUPAC_pos"]  = max_d_noamb if max_d_noamb else ""

            detail_parts = []
            for bn in comp_bns:
                ep = estados_por_bn.get(bn, "Absent")
                ep_es = {'identical': 'Identical', 'compatible': 'Compatible (IUPAC)',
                         'different': 'Different'}.get(ep, ep)
                nota_bn = f"{bn}: {ep_es}"
                d = dists_por_bn.get(bn, 0)
                if d:
                    nota_bn += f" (d={d})"
                iupac_p = noamb_por_bn.get(bn, 0)
                if ep == 'compatible' and iupac_p:
                    nota_bn += f" (IUPAC_pos={iupac_p})"
                if rc_por_bn.get(bn):
                    nota_bn += " [RC]"
                detail_parts.append(nota_bn)

            best = _best_barcode(diff_candidates) if diff_candidates else None
            if global_raw == 'different':
                if best:
                    row["Best_run"] = best[0]
                    detail_parts.append(
                        f"Best: {best[0]} (ambs={best[4]}, gaps={best[5]}, cov={best[3]})")
                else:
                    row["Best_run"] = "Tie"
                    detail_parts.append("Tie in quality")
            else:
                row["Best_run"] = ""

            row["Note"] = " | ".join(detail_parts)
            rows.append(row)
            self.notifyProgress.emit(int((prog_i + 1) / total * 100))

        has_iupac  = any(row.get("IUPAC_pos", "") for row in rows)
        has_length = any(row.get("Length", "")    for row in rows)
        headers = ["ID", "State", "Best_run", "Diff_bases"]
        if has_iupac:
            headers.append("IUPAC_pos")
        if has_length:
            headers.append("Length")
        headers += [f"{ref_bn}_len", f"{ref_bn}_cov",
                    f"{ref_bn}_ambs", f"{ref_bn}_gaps"]
        for bn in comp_bns:
            headers += [f"{bn}_len", f"{bn}_cov",
                        f"{bn}_ambs", f"{bn}_gaps", f"{bn}_dif",
                        f"{bn}_length"]
        headers.append("Note")

        _write_outputs(rows, headers, self.outdir,
                       all_bns=[ref_bn] + list(comp_bns),
                       ref_bn=ref_bn, seqs_store=seqs)
        self.taskFinished.emit(rows, headers)


class ComparePanel(QtWidgets.QWidget):
    # files, mode ("sets"|"ref"), ref_basename, outdir, extract_cfg (dict)
    compareRequested = QtCore.pyqtSignal(list, str, str, str, dict)

    def __init__(self, parent=None):
        super().__init__(parent)

        self._worker = None
        # Detached-but-maybe-running workers, kept referenced so they aren't GC'd
        # mid-run (which crashes Qt). Purged in _retire_worker.
        self._retired_workers: set = set()

        outer_layout = QtWidgets.QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)

        self._inner = QtWidgets.QWidget()
        self._layout = QtWidgets.QVBoxLayout(self._inner)
        self._layout.setContentsMargins(20, 20, 20, 8)
        self._layout.setSpacing(16)
        scroll.setWidget(self._inner)
        outer_layout.addWidget(scroll, 1)

        self._lbl_compare_title = make_label("Compare barcode sets", size=19, bold=True)
        self._lbl_compare_desc = make_label(
            "Compare multiFASTA files from different runs. "
            "The sample ID is extracted from the FASTA header based on a delimiter and occurrence. "
            "Use Advanced when files follow different header conventions: list several regex "
            "(tried in order) and normalize the result so the same sample collides across files.",
            color=TEXT_SEC
        )
        self._lbl_compare_desc.setWordWrap(True)
        self._layout.addWidget(self._lbl_compare_title)
        self._layout.addWidget(self._lbl_compare_desc)

        # ── ID extraction block ─────────────────────────────────────────────
        id_block = QtWidgets.QFrame()
        id_block_layout = QtWidgets.QVBoxLayout(id_block)
        id_block_layout.setContentsMargins(0, 2, 0, 2)
        id_block_layout.setSpacing(3)

        # Row 1: delimiter + occurrence + live preview
        simple_row = QtWidgets.QWidget()
        simple_layout = QtWidgets.QHBoxLayout(simple_row)
        simple_layout.setContentsMargins(0, 0, 0, 0)
        simple_layout.setSpacing(8)

        self._lbl_id_delim = make_label("ID delimiter:", color=TEXT_SEC)
        simple_layout.addWidget(self._lbl_id_delim, 0, QtCore.Qt.AlignVCenter)

        self._id_delim_combo = QtWidgets.QComboBox()
        self._id_delim_combo.setEditable(True)
        self._id_delim_combo.setFixedWidth(90)
        self._id_delim_combo.setToolTip(
            "Character (or string) that separates the sample ID from the rest of the header.\n"
            "Common choices: '_' for ONT outputs, '|' for BOLD/GenBank.\n"
            "Multi-character delimiters are also accepted (e.g. '_all.fa')."
        )
        for ch in ["_", "|", ";", " ", ".", "-"]:
            self._id_delim_combo.addItem(ch)
        self._id_delim_combo.setCurrentText("_")
        simple_layout.addWidget(self._id_delim_combo)

        self._lbl_id_occ = make_label("before occurrence:", color=TEXT_SEC)
        simple_layout.addWidget(self._lbl_id_occ, 0, QtCore.Qt.AlignVCenter)

        self._id_occ_combo = QtWidgets.QComboBox()
        self._id_occ_combo.setFixedWidth(62)
        self._id_occ_combo.setToolTip(
            "Which occurrence of the delimiter marks the end of the sample ID.\n"
            "Example with '_': '1st' → everything before the first '_'."
        )
        for suffix in ["1st", "2nd", "3rd", "4th", "5th"]:
            self._id_occ_combo.addItem(suffix)
        simple_layout.addWidget(self._id_occ_combo)

        self._id_delim_anyset_chk = QtWidgets.QCheckBox("any of these")
        self._id_delim_anyset_chk.setToolTip(
            "Treat the delimiter field as a SET of single characters and cut at\n"
            "whichever one appears first, instead of matching it literally.\n"
            "Use it when the same ID is followed by different separators across\n"
            "files, e.g. delimiter '_-' makes both 'BIOUG00045_all.fa' and\n"
            "'BIOUG00045-all.fa' yield 'BIOUG00045'."
        )
        simple_layout.addWidget(self._id_delim_anyset_chk)

        simple_layout.addSpacing(6)
        _arr = make_label("→", color=TEXT_SEC)
        simple_layout.addWidget(_arr, 0, QtCore.Qt.AlignVCenter)

        self._lbl_id_preview_head = make_label("Preview:", color=TEXT_SEC)
        simple_layout.addWidget(self._lbl_id_preview_head, 0, QtCore.Qt.AlignVCenter)

        self._lbl_id_preview = QtWidgets.QLabel("—")
        self._lbl_id_preview.setStyleSheet(
            f"color: {BLUE}; font-family: 'Courier New', monospace; font-weight: 600;"
            f" font-size: 19px; padding-top: 6px;"
        )
        self._lbl_id_preview.setAlignment(QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft)
        simple_layout.addWidget(self._lbl_id_preview, 1)
        id_block_layout.addWidget(simple_row)

        # Row 2: sample header preview card (shown after files are loaded)
        self._raw_header_frame = QtWidgets.QFrame()
        self._raw_header_frame.setObjectName("raw_header_card")
        self._raw_header_frame.setStyleSheet(f"""
            QFrame#raw_header_card {{
                background: {GRAY_BG};
                border: 1px solid {GRAY_LINE};
                border-radius: 6px;
                margin-top: 12px;
            }}
        """)
        rh_layout = QtWidgets.QVBoxLayout(self._raw_header_frame)
        rh_layout.setContentsMargins(10, 6, 10, 6)
        rh_layout.setSpacing(2)
        rh_title = QtWidgets.QLabel("Sample header")
        rh_title.setStyleSheet(f"font-size:14px; font-weight:600; color:{TEXT_HINT}; background:transparent;")
        rh_layout.addWidget(rh_title)
        self._lbl_raw_header = QtWidgets.QLabel("")
        self._lbl_raw_header.setStyleSheet(
            f"font-family:'Courier New',Consolas,monospace; font-size:15px;"
            f" color:{TEXT_SEC}; background:transparent;"
        )
        self._lbl_raw_header.setWordWrap(True)
        self._lbl_raw_header.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred
        )
        rh_layout.addWidget(self._lbl_raw_header)
        self._raw_header_frame.hide()
        id_block_layout.addWidget(self._raw_header_frame)

        # Row 3: Advanced (Regex) toggle
        adv_toggle_row = QtWidgets.QWidget()
        adv_toggle_layout = QtWidgets.QHBoxLayout(adv_toggle_row)
        adv_toggle_layout.setContentsMargins(0, 16, 0, 0)
        adv_toggle_layout.setSpacing(0)
        self._advanced_toggle_btn = QtWidgets.QPushButton("▶  Advanced (Regex)")
        self._advanced_toggle_btn.setObjectName("secondary_btn")
        self._advanced_toggle_btn.setCheckable(True)
        self._advanced_toggle_btn.setChecked(False)
        self._advanced_toggle_btn.setFixedHeight(40)
        self._advanced_toggle_btn.setFixedWidth(210)
        adv_toggle_layout.addWidget(self._advanced_toggle_btn)
        adv_toggle_layout.addStretch()
        id_block_layout.addWidget(adv_toggle_row)

        # Row 4: Advanced field (hidden by default) — multi-regex + normalize
        self._advanced_widget = QtWidgets.QWidget()
        adv_layout = QtWidgets.QVBoxLayout(self._advanced_widget)
        adv_layout.setContentsMargins(20, 0, 0, 0)
        adv_layout.setSpacing(8)

        regex_row = QtWidgets.QWidget()
        regex_layout = QtWidgets.QHBoxLayout(regex_row)
        regex_layout.setContentsMargins(0, 0, 0, 0)
        regex_layout.setSpacing(8)
        self._lbl_regex = make_label("Regex (one per line):", color=TEXT_SEC)
        self._lbl_regex.setAlignment(QtCore.Qt.AlignTop)
        regex_layout.addWidget(self._lbl_regex, 0, QtCore.Qt.AlignTop)
        self._regex_edit = QtWidgets.QPlainTextEdit()
        self._regex_edit.setFixedWidth(600)
        self._regex_edit.setFixedHeight(70)
        self._regex_edit.setPlaceholderText(
            "One regex per line — tried in order until one matches.\n"
            r"e.g. ^(\S+?)_all\.fa,   ^([^|]+)\|"
        )
        self._regex_edit.setToolTip(
            "One regular expression per line. Each FASTA header is tried against\n"
            "every line in order until one matches, so files with different header\n"
            "conventions can share a single config (Option B).\n"
            "If a regex has a capturing group, group 1 is used as the ID;\n"
            "otherwise the full match is used.\n"
            "Leave empty to use the positional mode above."
        )
        regex_layout.addWidget(self._regex_edit)
        self._lbl_regex_status = make_label("", color=TEXT_SEC)
        self._lbl_regex_status.setFixedWidth(140)
        self._lbl_regex_status.setAlignment(QtCore.Qt.AlignTop)
        regex_layout.addWidget(self._lbl_regex_status, 0, QtCore.Qt.AlignTop)
        regex_layout.addStretch()
        adv_layout.addWidget(regex_row)

        norm_row = QtWidgets.QWidget()
        norm_layout = QtWidgets.QHBoxLayout(norm_row)
        norm_layout.setContentsMargins(0, 0, 0, 0)
        norm_layout.setSpacing(8)
        self._lbl_norm = make_label("Normalize ID:", color=TEXT_SEC)
        norm_layout.addWidget(self._lbl_norm, 0, QtCore.Qt.AlignVCenter)
        self._norm_lower_chk = QtWidgets.QCheckBox("lowercase")
        self._norm_lower_chk.setToolTip("Casefold the extracted ID (BIN-A → bin-a).")
        norm_layout.addWidget(self._norm_lower_chk)
        self._norm_zeros_chk = QtWidgets.QCheckBox("strip leading zeros")
        self._norm_zeros_chk.setToolTip(
            "Drop leading zeros inside numeric runs (BIOUG00045 → BIOUG45)."
        )
        norm_layout.addWidget(self._norm_zeros_chk)
        self._lbl_norm_strip = make_label("strip (regex):", color=TEXT_SEC)
        norm_layout.addWidget(self._lbl_norm_strip, 0, QtCore.Qt.AlignVCenter)
        self._norm_strip_edit = QtWidgets.QLineEdit()
        self._norm_strip_edit.setFixedWidth(180)
        self._norm_strip_edit.setPlaceholderText(r"e.g. _S\d+$  or  -RUN\d+")
        self._norm_strip_edit.setToolTip(
            "Substrings matching this regex are removed from the ID after\n"
            "extraction — useful to drop run/replicate suffixes so the same\n"
            "sample collides across files (Option C). Leave empty to disable."
        )
        norm_layout.addWidget(self._norm_strip_edit)
        norm_layout.addStretch()
        adv_layout.addWidget(norm_row)

        self._advanced_widget.hide()
        id_block_layout.addWidget(self._advanced_widget)

        self._layout.addWidget(id_block)

        # Connect ID extraction signals
        self._id_delim_combo.currentTextChanged.connect(self._update_id_preview)
        self._id_occ_combo.currentIndexChanged.connect(self._update_id_preview)
        self._id_delim_anyset_chk.toggled.connect(self._update_id_preview)
        self._advanced_toggle_btn.toggled.connect(self._on_advanced_toggled)
        self._regex_edit.textChanged.connect(self._on_regex_changed)
        self._norm_lower_chk.toggled.connect(self._update_id_preview)
        self._norm_zeros_chk.toggled.connect(self._update_id_preview)
        self._norm_strip_edit.textChanged.connect(self._update_id_preview)

        self._mode_box = QtWidgets.QGroupBox("Comparison mode")
        self._mode_box.setStyleSheet("QGroupBox { font-weight:600; color:#1A1A2E; }")
        mb_layout = QtWidgets.QVBoxLayout(self._mode_box)
        self._radio_sets = QtWidgets.QRadioButton(
            "Compare sets with each other — identical /IUPAC compatible /different /unique"
        )
        self._radio_sets_src = "Compare sets with each other — identical /IUPAC compatible /different /unique"
        self._radio_sets.setChecked(True)
        self._radio_ref = QtWidgets.QRadioButton(
            "Compare against reference — a file acts as a reference"
        )
        self._radio_ref_src = "Compare against reference — a file acts as a reference"
        mb_layout.addWidget(self._radio_sets)
        mb_layout.addWidget(self._radio_ref)
        self._layout.addWidget(self._mode_box)

        self._ref_widget = QtWidgets.QWidget()
        ref_layout = QtWidgets.QHBoxLayout(self._ref_widget)
        ref_layout.setContentsMargins(0, 0, 0, 0)
        self._lbl_ref_file = make_label("Reference file:", color=TEXT_SEC)
        ref_layout.addWidget(self._lbl_ref_file)
        self._ref_combo = QtWidgets.QComboBox()
        self._ref_combo.setMinimumWidth(320)
        ref_layout.addWidget(self._ref_combo)
        ref_layout.addStretch()
        self._ref_widget.hide()
        self._layout.addWidget(self._ref_widget)

        self._radio_ref.toggled.connect(self._on_mode_toggled)

        # Output folder state — the save location is chosen in the dialog opened
        # by "Start comparison" (_ask_output_dir), so no card is shown here.
        self._last_outdir = ""

        self._drop = MultiDropZone()
        self._drop.filesDropped.connect(self._on_files)
        self._layout.addWidget(self._drop)

        self._comp_bar = QtWidgets.QProgressBar()
        self._comp_bar.setRange(0, 100)
        self._comp_bar.setValue(0)
        self._comp_bar.hide()
        self._layout.addWidget(self._comp_bar)

        self._result_lbl = make_label("", color=TEXT_SEC)
        self._result_lbl.setWordWrap(True)
        self._layout.addWidget(self._result_lbl)

        self._layout.addStretch()
        self._results_win = None

        footer = QtWidgets.QWidget()
        footer.setObjectName("compare_footer")
        footer.setStyleSheet(f"""
            QWidget#compare_footer {{
                background: {GRAY_CARD};
                border-top: 1px solid {GRAY_LINE};
            }}
        """)
        fl = QtWidgets.QHBoxLayout(footer)
        fl.setContentsMargins(20, 10, 20, 10)

        self._clear_btn = QtWidgets.QPushButton("Clear")
        self._clear_btn.setObjectName("danger_btn")
        self._clear_btn.setFixedHeight(44)
        self._clear_btn.setFixedWidth(140)
        self._clear_btn.clicked.connect(self._reset)
        fl.addWidget(self._clear_btn)

        self._view_btn = QtWidgets.QPushButton("Show table  ↗")
        self._view_btn.setObjectName("secondary_btn")
        self._view_btn.setFixedHeight(44)
        self._view_btn.hide()
        self._view_btn.clicked.connect(self._open_results_win)
        fl.addWidget(self._view_btn)

        self._open_folder_btn = QtWidgets.QPushButton("Open folder  📂")
        self._open_folder_btn.setObjectName("secondary_btn")
        self._open_folder_btn.setFixedHeight(44)
        self._open_folder_btn.hide()
        self._open_folder_btn.clicked.connect(self._open_output_folder)
        fl.addWidget(self._open_folder_btn)

        self._open_excel_btn = QtWidgets.QPushButton("Open Excel  📊")
        self._open_excel_btn.setObjectName("secondary_btn")
        self._open_excel_btn.setFixedHeight(44)
        self._open_excel_btn.hide()
        self._open_excel_btn.clicked.connect(self._open_summary_excel)
        fl.addWidget(self._open_excel_btn)

        fl.addStretch()

        self._compare_btn = QtWidgets.QPushButton("Start comparison  →")
        self._compare_btn.setObjectName("primary_btn")
        self._compare_btn.setFixedHeight(44)
        self._compare_btn.setFixedWidth(300)
        self._compare_btn.setEnabled(False)
        self._compare_btn.clicked.connect(self._emit_compare)
        self._compare_btn.setStyleSheet(
            f"QPushButton {{ background-color: {GRAY_LINE}; color: {TEXT_HINT}; border:none; "
            f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
        )
        fl.addWidget(self._compare_btn)
        outer_layout.addWidget(footer)

        self._compare_btn_ref = self._compare_btn
        self.compareRequested.connect(self._start_compare)

    def retranslateUi(self):
        ctx = "ComparePanel"
        self._lbl_compare_title.setText(_tr(ctx, "Compare barcode sets"))
        self._lbl_compare_desc.setText(_tr(ctx,
            "Compare multiFASTA files from different runs. "
            "The sample ID is extracted from the FASTA header based on a delimiter and occurrence. "
            "Use Advanced when files follow different header conventions: list several regex "
            "(tried in order) and normalize the result so the same sample collides across files."))
        self._lbl_id_delim.setText(_tr(ctx, "ID delimiter:"))
        self._lbl_id_occ.setText(_tr(ctx, "before occurrence:"))
        self._id_delim_anyset_chk.setText(_tr(ctx, "any of these"))
        self._lbl_id_preview_head.setText(_tr(ctx, "Preview:"))
        arrow = "▼  " if self._advanced_toggle_btn.isChecked() else "▶  "
        self._advanced_toggle_btn.setText(arrow + _tr(ctx, "Advanced (Regex)"))
        self._lbl_regex.setText(_tr(ctx, "Regex (one per line):"))
        self._lbl_norm.setText(_tr(ctx, "Normalize ID:"))
        self._norm_lower_chk.setText(_tr(ctx, "lowercase"))
        self._norm_zeros_chk.setText(_tr(ctx, "strip leading zeros"))
        self._lbl_norm_strip.setText(_tr(ctx, "strip (regex):"))
        self._mode_box.setTitle(_tr(ctx, "Comparison mode"))
        self._radio_sets.setText(_tr(ctx, self._radio_sets_src))
        self._radio_ref.setText(_tr(ctx, self._radio_ref_src))
        self._lbl_ref_file.setText(_tr(ctx, "Reference file:"))
        self._clear_btn.setText(_tr(ctx, "Clear"))
        self._view_btn.setText(_tr(ctx, "Show table  ↗"))
        self._open_folder_btn.setText(_tr(ctx, "Open folder  📂"))
        self._open_excel_btn.setText(_tr(ctx, "Open Excel  📊"))
        self._compare_btn.setText(_tr(ctx, "Start comparison  →"))
        self._drop.retranslateUi()

    def changeEvent(self, event):
        if event.type() == QtCore.QEvent.LanguageChange:
            self.retranslateUi()
        super().changeEvent(event)

    def _retire_worker(self):
        """Detach the current compare worker: disconnect its slots so stale results
        can't reach the UI, and keep a reference until its thread exits so it is never
        garbage-collected while still running (which crashes Qt). These workers are
        CPU-bound (edlib) and have no stop(), so they simply run to completion."""
        self._retired_workers = {w for w in self._retired_workers if w.isRunning()}
        w = self._worker
        self._worker = None
        if w is None:
            return
        for sig in (w.notifyProgress, w.taskFinished):
            try:
                sig.disconnect()
            except (TypeError, RuntimeError):
                pass
        if w.isRunning():
            self._retired_workers.add(w)

    def _on_mode_toggled(self, ref_active):
        self._ref_widget.setVisible(ref_active)

    def _on_advanced_toggled(self, checked: bool):
        self._advanced_widget.setVisible(checked)
        ctx = "ComparePanel"
        arrow = "▼  " if checked else "▶  "
        self._advanced_toggle_btn.setText(arrow + _tr(ctx, "Advanced (Regex)"))
        self._update_id_preview()

    def _on_regex_changed(self, *_):
        lines = [ln.strip() for ln in self._regex_edit.toPlainText().splitlines()
                 if ln.strip()]
        if not lines:
            self._lbl_regex_status.setText("")
            self._lbl_regex_status.setStyleSheet("")
        else:
            bad = None
            for ln in lines:
                try:
                    re.compile(ln)
                except re.error as e:
                    bad = (ln, e)
                    break
            if bad is None:
                self._lbl_regex_status.setText(f"✓ {len(lines)} valid")
                self._lbl_regex_status.setStyleSheet("color: #16A34A;")
            else:
                self._lbl_regex_status.setText(f"✗ {bad[1].msg}")
                self._lbl_regex_status.setStyleSheet("color: #DC2626;")
        self._update_id_preview()

    def _get_extract_cfg(self) -> dict:
        delim = self._id_delim_combo.currentText()
        occ = self._id_occ_combo.currentIndex() + 1
        if delim:
            kind = "posany" if self._id_delim_anyset_chk.isChecked() else "pos"
            id_patterns = [f"{kind}:{delim}:{occ}"]
        else:
            id_patterns = []

        regex_patterns: list = []
        normalize: dict = {}
        if self._advanced_toggle_btn.isChecked():
            regex_patterns = [ln.strip()
                              for ln in self._regex_edit.toPlainText().splitlines()
                              if ln.strip()]
            normalize = {
                "lowercase":   self._norm_lower_chk.isChecked(),
                "strip_zeros": self._norm_zeros_chk.isChecked(),
                "strip_regex": self._norm_strip_edit.text().strip(),
            }
        return {
            "id_patterns":    id_patterns,
            "regex_patterns": regex_patterns,
            "normalize":      normalize,
        }

    def _update_id_preview(self, *_):
        files = getattr(self._drop, 'files', [])
        if not files:
            self._lbl_id_preview.setText("—")
            self._lbl_raw_header.setText("")
            self._raw_header_frame.hide()
            return
        first_header = None
        try:
            with open(files[0], encoding="utf-8", errors="replace") as fh:
                for line in fh:
                    stripped = line.strip()
                    if stripped.startswith(">"):
                        first_header = stripped
                        break
        except Exception:
            pass
        if not first_header:
            self._lbl_id_preview.setText("—")
            self._lbl_raw_header.setText("")
            self._raw_header_frame.hide()
            return
        raw_display = first_header[:100] + ("…" if len(first_header) > 100 else "")
        self._lbl_raw_header.setText(raw_display)
        self._raw_header_frame.show()
        cfg = self._get_extract_cfg()
        try:
            sid, *_ = _parse_fasta_header(first_header, cfg)
            self._lbl_id_preview.setText(sid if sid else "—")
        except Exception:
            self._lbl_id_preview.setText("—")

    def _emit_compare(self):
        mode = "ref" if self._radio_ref.isChecked() else "sets"
        ref_path = (self._ref_combo.currentData() or "") if mode == "ref" else ""
        extract_cfg = self._get_extract_cfg()
        # The output folder is resolved in _ask_output_dir (opened from _start_compare),
        # so no custom folder is passed from the panel.
        self.compareRequested.emit(
            self._drop.files, mode, ref_path, "", extract_cfg
        )

    def _on_files(self, paths):
        self._ref_combo.clear()
        labels = _make_unique_labels(paths) if paths else []
        for lbl, p in zip(labels, paths):
            self._ref_combo.addItem(lbl, p)
        enabled = len(paths) >= 2
        self._compare_btn_ref.setEnabled(enabled)
        if enabled:
            self._compare_btn_ref.setStyleSheet(
                f"QPushButton {{ background-color: {BLUE}; color: white; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
                f"QPushButton:hover {{ background-color: #0C4A82; }}"
            )
        else:
            self._compare_btn_ref.setStyleSheet(
                f"QPushButton {{ background-color: {GRAY_LINE}; color: {TEXT_HINT}; border:none; "
                f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
            )
        self._update_id_preview()

    def _ask_output_dir(self, auto_dir: str, folder_name: str) -> Optional[str]:
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Output folder")
        dlg.setMinimumWidth(480)
        dlg.setStyleSheet(f"""
            QDialog {{ background-color: {GRAY_CARD}; }}
            QLabel {{ color: {TEXT_PRI}; background-color: transparent; }}
            QRadioButton {{
                color: {TEXT_PRI}; background-color: transparent;
                font-size: 15px; padding: 6px 0;
            }}
            QRadioButton::indicator {{ width: 16px; height: 16px; }}
            QPushButton {{
                border-radius: 8px; padding: 8px 20px;
                font-size: 15px; font-weight: 500;
            }}
            #dlg_ok_btn {{ background-color: {BLUE}; color: white; border: none; }}
            #dlg_ok_btn:hover {{ background-color: #0C4A82; }}
            #dlg_cancel_btn {{
                background-color: transparent; color: {BLUE};
                border: 1px solid {BLUE};
            }}
            #dlg_cancel_btn:hover {{ background-color: {BLUE_LIGHT}; }}
        """)
        vlay = QtWidgets.QVBoxLayout(dlg)
        vlay.setSpacing(16)
        vlay.setContentsMargins(24, 24, 24, 20)

        title_lbl = QtWidgets.QLabel("Where to save the results?")
        title_lbl.setStyleSheet(f"font-size:17px; font-weight:700; color:{TEXT_PRI};")
        vlay.addWidget(title_lbl)

        radio_auto = QtWidgets.QRadioButton(
            f"Automatic folder (recommended)\n  …/output/{folder_name}/"
        )
        radio_auto.setChecked(True)
        radio_custom = QtWidgets.QRadioButton("Select folder manually")
        vlay.addWidget(radio_auto)
        vlay.addWidget(radio_custom)
        vlay.addSpacing(8)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_cancel.setObjectName("dlg_cancel_btn")
        btn_cancel.setFixedHeight(38)
        btn_ok = QtWidgets.QPushButton("Run")
        btn_ok.setObjectName("dlg_ok_btn")
        btn_ok.setFixedHeight(38)
        btn_ok.setDefault(True)
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addSpacing(8)
        btn_row.addWidget(btn_ok)
        vlay.addLayout(btn_row)

        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return None

        if radio_auto.isChecked():
            return auto_dir
        parent_dir = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select output folder"
        )
        if not parent_dir:
            return None
        return os.path.join(parent_dir, folder_name)

    def _start_compare(self, files, mode, ref_path, custom_outdir, extract_cfg):
        ts          = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"ont-barcoder_{ts}_comp"
        auto_dir    = os.path.join(_get_base_dir(), "output", folder_name)

        if custom_outdir:
            outdir = custom_outdir
        else:
            outdir = self._ask_output_dir(auto_dir, folder_name)
            if outdir is None:
                return

        self._last_outdir = outdir

        self._compare_btn_ref.setEnabled(False)
        self._comp_bar.setRange(0, 100)
        self._comp_bar.setValue(0)
        self._comp_bar.show()
        self._result_lbl.setText("Running comparison…")
        self._view_btn.hide()
        self._open_folder_btn.hide()
        self._open_excel_btn.hide()

        self._retire_worker()
        if mode == "ref":
            self._worker = _PairCompareWorker(
                files, ref_path, outdir, extract_cfg)
        else:
            self._worker = _CompareWorker(files, outdir, extract_cfg)

        self._worker.notifyProgress.connect(self.update_progress)
        self._worker.taskFinished.connect(self.show_results)
        self._worker.start()

    @QtCore.pyqtSlot(int)
    def update_progress(self, n):
        if self._comp_bar.maximum() == 0:
            self._comp_bar.setRange(0, 100)
        self._comp_bar.show()
        self._comp_bar.setValue(n)

    @QtCore.pyqtSlot(list, list)
    def show_results(self, rows, headers):
        self._comp_bar.hide()
        self._comp_bar.setValue(0)

        n_identical = sum(1 for r in rows if r.get("State") == "Identical")
        n_compat    = sum(1 for r in rows if r.get("State") == "Compatible (IUPAC)")
        n_diff      = sum(1 for r in rows if r.get("State") == "Different")
        n_only      = sum(1 for r in rows if r.get("State", "").startswith("Unique"))
        n_only_ref  = sum(1 for r in rows if r.get("State") == "Only in reference")
        n_no_ref    = sum(1 for r in rows if r.get("State") == "No reference")

        ctx = "ComparePanel"
        summary = (
            f"<b>{_tr(ctx, 'Comparison completed.')}</b> {_tr(ctx, 'Total IDs')}: {len(rows)} &nbsp;|&nbsp; "
            f"🟢 {_tr(ctx, 'Identical')}: {n_identical}"
        )
        if n_compat:
            summary += f" &nbsp;|&nbsp; 🔵 {_tr(ctx, 'Compatibles (IUPAC)')}: {n_compat}"
        if n_diff:
            summary += f" &nbsp;|&nbsp; 🔴 {_tr(ctx, 'Different')}: {n_diff}"
        if n_only:
            summary += f" &nbsp;|&nbsp; ⚪ {_tr(ctx, 'Unique')}: {n_only}"
        if n_only_ref:
            summary += f" &nbsp;|&nbsp; 📌 {_tr(ctx, 'Only in reference')}: {n_only_ref}"
        if n_no_ref:
            summary += f" &nbsp;|&nbsp; ❓ {_tr(ctx, 'No reference')}: {n_no_ref}"

        outdir_shown = self._last_outdir
        if outdir_shown:
            summary += f"<br>📁 {_tr(ctx, 'Output in')}: <i>{outdir_shown}</i>"
        self._result_lbl.setText(summary)

        self._compare_btn_ref.setEnabled(True)
        self._open_folder_btn.show()
        self._open_excel_btn.show()

        if self._results_win is not None:
            try:
                self._results_win.close()
            except Exception:
                pass
        self._results_win = _CompareResultsWindow(rows, headers, outdir_shown, self)
        self._results_win.show()
        self._view_btn.show()

    def _open_results_win(self):
        if self._results_win is not None:
            self._results_win.show()
            self._results_win.raise_()

    def _open_output_folder(self):
        path = self._last_outdir
        if path and os.path.isdir(path):
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))

    def _open_summary_excel(self):
        path = self._last_outdir
        if path:
            xlsx = os.path.join(path, "summary.xlsx")
            if os.path.isfile(xlsx):
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(xlsx))

    def _reset(self):
        self._retire_worker()
        if self._results_win is not None:
            try:
                self._results_win.close()
            except Exception:
                pass
            self._results_win = None

        self._drop.clear()
        self._comp_bar.hide()
        self._comp_bar.setValue(0)
        self._comp_bar.setRange(0, 100)
        self._result_lbl.setText("")

        # Reset output folder state
        self._last_outdir = ""

        self._view_btn.hide()
        self._open_folder_btn.hide()
        self._open_excel_btn.hide()
        self._compare_btn_ref.setEnabled(False)
        self._compare_btn_ref.setStyleSheet(
            f"QPushButton {{ background-color: {GRAY_LINE}; color: {TEXT_HINT}; border:none; "
            f"border-radius:8px; padding:9px 20px; font-size:18px; font-weight:500; }}"
        )

        # Reset ID extraction controls
        self._id_delim_combo.setCurrentText("_")
        self._id_occ_combo.setCurrentIndex(0)
        self._id_delim_anyset_chk.setChecked(False)
        self._advanced_toggle_btn.setChecked(False)
        self._advanced_widget.hide()
        self._regex_edit.setPlainText("")
        self._lbl_regex_status.setText("")
        self._norm_lower_chk.setChecked(False)
        self._norm_zeros_chk.setChecked(False)
        self._norm_strip_edit.setText("")
        self._lbl_id_preview.setText("—")
        self._lbl_raw_header.setText("")
        self._raw_header_frame.hide()

        self._radio_sets.setChecked(True)
        self._ref_widget.hide()
        self._ref_combo.clear()


def main():
    multiprocessing.freeze_support()  # required for frozen (PyInstaller) multiprocessing

    # On Windows, give the app its own taskbar identity so the taskbar shows our
    # icon instead of the generic Python one. Must run before any window is shown.
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "BarcodeSuite.app"
            )
        except Exception:
            pass

    app = QtWidgets.QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)

    # Window + taskbar icon (icon.ico sits next to the .exe / this script).
    _icon_path = os.path.join(_get_base_dir(), "icon.ico")
    if os.path.exists(_icon_path):
        _app_icon = QtGui.QIcon(_icon_path)
        app.setWindowIcon(_app_icon)

    settings = QtCore.QSettings("BarcodeSuite")

    win = QtWidgets.QMainWindow()
    win.setWindowTitle('BarcodeSuite')
    if os.path.exists(_icon_path):
        win.setWindowIcon(_app_icon)
    win.setMinimumSize(1450, 1020)  # <-- tamaño mínimo (ancho suficiente para mostrar todas las pestañas)
    central = QtWidgets.QWidget()
    win.setCentralWidget(central)
    layout = QtWidgets.QVBoxLayout(central)
    layout.setContentsMargins(0, 0, 0, 0)
    
    tabs = QtWidgets.QTabWidget()
    tabs.tabBar().setExpanding(True)
    tabs.setDocumentMode(True)  # Modo más limpio
    # Opcional: establecer tamaño mínimo para cada pestaña
    tabs.tabBar().setMinimumWidth(200)
    tabs.tabBar().setExpanding(True)
    
    tabs.addTab(FastqInspectorPanel(), 'FASTQ Inspector')
    tabs.addTab(FastaToolsPanel(), 'FASTA Tools')
    tabs.addTab(ComparePanel(), 'FASTA Compare')
    tabs.addTab(BlastPanel(), 'BLAST')
    tabs.addTab(GenbankBatchPanel(), 'GenBank Batch')
    tabs.addTab(BoldFormatPanel(), 'BOLD Formatter')
    layout.addWidget(tabs)

    geom = settings.value("mainwindow/geometry")
    if geom:
        win.restoreGeometry(geom)
    else:
        win.resize(1450, 1020)

    app.aboutToQuit.connect(
        lambda: settings.setValue("mainwindow/geometry", win.saveGeometry())
    )

    win.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
