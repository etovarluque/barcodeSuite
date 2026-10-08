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


def _install_package(module, package):
    """Install `package` with pip if `module` cannot be imported."""
    if importlib.util.find_spec(module) is None:
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
    # (import name, pip package name)
    _REQUIRED_PACKAGES = [
        ("PyQt5", "PyQt5"), ("edlib", "edlib"), ("xlsxwriter", "xlsxwriter"),
        ("openpyxl", "openpyxl"), ("Bio", "biopython"),
    ]
    _check_and_install_pip()
    for _mod, _pkg in _REQUIRED_PACKAGES:
        _install_package(_mod, _pkg)

import os
import time
import datetime
import multiprocessing
import warnings

# Suppress non-relevant warnings from Biopython (partial codons, etc.)
warnings.filterwarnings("ignore", message="Partial codon")
warnings.filterwarnings("ignore", category=UserWarning, module="Bio")

from PyQt5 import QtCore, QtGui, QtWidgets

import suite_ui
from suite_ui import Sidebar, apply_theme, current_theme, set_dark_titlebar

# Utility panels are shared verbatim with ONTbarcoder3 (_utilities/), so they
# can be re-synced by copying the modules over. GenBank Batch is Suite-only.
from _utilities.shared import (
    STYLESHEET, BLUE, BLUE_LIGHT, GRAY_CARD, TEXT_PRI, _get_base_dir, _tr,
)
from _utilities.fastq_inspector import FastqInspectorPanel
from _utilities.fasta_tools import FastaToolsPanel
from _utilities.compare_panel import ComparePanel, _CompareWorker, _PairCompareWorker
from _utilities import compare_panel as _compare_panel
from _utilities.blast_panel import BlastPanel, _BlastWorker, _BlastFileWorker
from _utilities.genbank_batch import GenbankBatchPanel
from _utilities.bold_formatter import BoldFormatterPanel

__version__ = "2.1.1"
APP_NAME = "BarcodeSuite"

# Sidebar layout: (section, [(tool key, label, icon), ...])
_TOOL_SECTIONS = [
    ("Reads",          [("fastq", "FASTQ Inspector", "fastq")]),
    ("Sequences",      [("fasta", "FASTA Tools", "fasta"),
                        ("compare", "FASTA Compare", "compare")]),
    ("Identification", [("blast", "BLAST", "blast")]),
    ("Databases",      [("genbank", "GenBank Batch", "genbank"),
                        ("bold", "BOLD Formatter", "bold")]),
]


def _settings() -> QtCore.QSettings:
    return QtCore.QSettings(APP_NAME)


class MainWindow(QtWidgets.QMainWindow):
    """Sidebar host for the utility panels. Compare and BLAST
    only emit requests; the workers are started and wired here (same glue as
    ONTbarcoder3's MainWindow)."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {__version__}")
        self.setMinimumSize(1200, 800)

        self._panel_fastq_inspector = FastqInspectorPanel()
        self._panel_fasta_tools     = FastaToolsPanel()
        self._panel_compare         = ComparePanel()
        self._panel_blast           = BlastPanel()
        self._panel_genbank         = GenbankBatchPanel()
        self._panel_bold_formatter  = BoldFormatterPanel()

        self._panels = {
            "fastq":    self._panel_fastq_inspector,
            "fasta":    self._panel_fasta_tools,
            "compare":  self._panel_compare,
            "blast":    self._panel_blast,
            "genbank":  self._panel_genbank,
            "bold":     self._panel_bold_formatter,
        }
        self._stack = QtWidgets.QStackedWidget()
        for panel in self._panels.values():
            self._stack.addWidget(panel)

        self._sidebar = Sidebar(_TOOL_SECTIONS, __version__)
        content = QtWidgets.QWidget()
        content.setObjectName("suite_content")
        cl = QtWidgets.QVBoxLayout(content)
        cl.setContentsMargins(8, 8, 8, 0)
        cl.addWidget(self._stack)

        central = QtWidgets.QWidget()
        root = QtWidgets.QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self._sidebar)
        root.addWidget(content, 1)
        self.setCentralWidget(central)

        # Restore the last session's choices (theme is applied in main()).
        cfg = _settings()
        self._sidebar.set_collapsed(cfg.value("ui/sidebar_collapsed", False, type=bool))
        self._switch_tool(cfg.value("ui/current_tool", "fastq", type=str))

        self._sidebar.toolSelected.connect(self._switch_tool)
        self._sidebar.themeToggled.connect(self._toggle_theme)
        self._sidebar.collapsedChanged.connect(
            lambda c: _settings().setValue("ui/sidebar_collapsed", c))
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+B"), self, activated=lambda:
                            self._sidebar.set_collapsed(not self._sidebar.is_collapsed(), True))

        self._panel_compare.compareRequested.connect(self._start_comparison)
        self._panel_blast.blastRequested.connect(self._start_blast)
        self._panel_blast.resumeRequested.connect(self._resume_blast)
        self._panel_blast.blastFileRequested.connect(self._start_blast_file)
        # Best Sequence is not part of the Suite (it is tied to ONTbarcoder's
        # multi-run workflow): drop BLAST's hand-off button. blast_panel.py is
        # shared verbatim, so the button is neutralised here instead.
        btn = self._panel_blast._send_best_btn
        btn.hide()
        btn.show = lambda: None
        desc = self._panel_blast._lbl_desc
        desc.setText(desc.text().replace(
            ", ready to use as a pair in Best Sequence.", "."))

    # ── Navigation / appearance ──────────────────────────────────────────────

    def _switch_tool(self, key: str):
        if key not in self._panels:
            key = "fastq"
        self._stack.setCurrentWidget(self._panels[key])
        self._sidebar.select(key)
        _settings().setValue("ui/current_tool", key)

    def _toggle_theme(self):
        theme = "light" if current_theme() == "dark" else "dark"
        apply_theme(QtWidgets.QApplication.instance(), theme, STYLESHEET)
        self._sidebar.refresh()
        _settings().setValue("ui/theme", theme)

    def showEvent(self, event):
        super().showEvent(event)
        set_dark_titlebar(self, current_theme() == "dark")

    # ── Output folder ────────────────────────────────────────────────────────

    def _ask_outdir(self, suffix: str) -> str:
        """Ask where to save a run's results. Returns the (not yet created)
        output folder, or '' if the user cancelled."""
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"ont-barcoder_{ts}_{suffix}"
        default_out = os.path.join(_get_base_dir(), "output", folder_name)

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle(_tr("MainWindow", "Output folder"))
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
        radio_default = QtWidgets.QRadioButton(
            f"Automatic folder (recommended)\n  …/output/{folder_name}")
        radio_default.setChecked(True)
        radio_custom = QtWidgets.QRadioButton("Select folder manually")
        vlay.addWidget(radio_default)
        vlay.addWidget(radio_custom)
        vlay.addSpacing(8)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_cancel.setObjectName("dlg_cancel_btn")
        btn_cancel.setFixedHeight(38)
        btn_ok = QtWidgets.QPushButton("Continue")
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
            return ""
        if radio_default.isChecked():
            return default_out
        parent_dir = QtWidgets.QFileDialog.getExistingDirectory(
            self, _tr("MainWindow", "Select the output folder"))
        return os.path.join(parent_dir, folder_name) if parent_dir else ""

    # ── Worker lifecycle ─────────────────────────────────────────────────────

    def _retire_worker(self, w):
        """Stop a QThread cooperatively, never terminate()-ing it. One still
        inside a (time-capped) socket call is kept referenced until it ends."""
        try:
            if not w.isRunning():
                return
            w.stop()
            if w.wait(3000):
                return
        except Exception:
            return
        retired = self.__dict__.setdefault("_retired_blast_workers", [])
        retired.append(w)
        w.finished.connect(lambda w=w: retired.remove(w) if w in retired else None)

    def _drop_worker(self, attr: str):
        """Disconnect and retire the worker stored in self.<attr>, if any."""
        w = getattr(self, attr, None)
        if w is None:
            return
        for sig in ("statusUpdated", "progressUpdated", "taskFinished", "taskError"):
            try:
                getattr(w, sig).disconnect()
            except (RuntimeError, TypeError, AttributeError):
                pass
        self._retire_worker(w)
        setattr(self, attr, None)

    # ── FASTA Compare ────────────────────────────────────────────────────────

    def _start_comparison(self, file_list: list, mode: str, ref_path: str,
                          custom_outdir: str, extract_cfg: dict = None):
        pnl = self._panel_compare
        if len(file_list) < 2:
            pnl._result_lbl.setText("At least 2 files are needed to compare.")
            return
        if mode == "ref" and not ref_path:
            pnl._result_lbl.setText("Select a reference file.")
            return

        # A folder chosen with "Change…" in the panel is used as is.
        if custom_outdir and os.path.isdir(custom_outdir):
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            outdir = os.path.join(custom_outdir, f"ont-barcoder_{ts}_comp")
        else:
            outdir = self._ask_outdir("comp")
            if not outdir:
                return
        try:
            os.makedirs(outdir, exist_ok=True)
        except Exception as e:
            pnl._result_lbl.setText(f"Error creating output folder: {e}")
            return

        pnl.set_outdir(outdir)
        pnl._result_lbl.setText(_tr("ComparePanel", "Comparison in progress…"))
        pnl._comp_bar.setRange(0, 0)
        pnl._comp_bar.show()
        pnl._compare_btn_ref.setEnabled(False)

        if mode == "ref":
            self.comp_worker = _PairCompareWorker(file_list, ref_path, outdir, extract_cfg)
        else:
            self.comp_worker = _CompareWorker(file_list, outdir, extract_cfg)
        self.comp_worker.notifyProgress.connect(pnl.update_progress)
        self.comp_worker.writeErrors.connect(pnl.show_write_errors)
        self.comp_worker.taskFinished.connect(pnl.show_results)
        self.comp_worker.taskError.connect(pnl.on_compare_error)
        self.comp_worker.start()

    # ── BLAST ────────────────────────────────────────────────────────────────

    def _blast_file_tab_idle(self) -> bool:
        # Both BLAST tabs keep their own NCBI rate limiter; running them at once
        # would double the request rate against the same IP (429s).
        other = getattr(self, "blast_file_worker", None)
        if other is not None and other.isRunning():
            self._panel_blast.on_error(
                "The 'BLAST web results' tab is still running. Wait for it to "
                "finish, or click its Stop button, before starting a new BLAST search.")
            return False
        return True

    def _start_blast(self, files: list, cfg: dict):
        if not self._blast_file_tab_idle():
            return
        outdir = self._ask_outdir("blast")
        if not outdir:
            return
        try:
            os.makedirs(outdir, exist_ok=True)
        except Exception as e:
            self._panel_blast.on_error(f"Could not create output folder: {e}")
            return
        cfg["outdir"] = outdir
        self._launch_blast_worker(files, cfg)

    def _resume_blast(self, path: str, cfg: dict):
        """Continue a stopped BLAST run in its own folder and files."""
        if not self._blast_file_tab_idle():
            return
        cfg["resume"] = path
        cfg["outdir"] = os.path.dirname(os.path.abspath(path))
        self._launch_blast_worker([], cfg)

    def _launch_blast_worker(self, files: list, cfg: dict):
        pnl = self._panel_blast
        pnl.set_running(True)
        pnl.update_status("result", f"Output    │ {cfg['outdir']}")

        try:
            pnl.stopRequested.disconnect()
        except (RuntimeError, TypeError):
            pass
        self._drop_worker("blast_worker")

        self.blast_worker = _BlastWorker(files, cfg)
        pnl.stopRequested.connect(self.blast_worker.stop)
        self.blast_worker.statusUpdated.connect(pnl.update_status)
        self.blast_worker.progressUpdated.connect(pnl.set_progress)
        self.blast_worker.taskFinished.connect(pnl.on_finished)
        self.blast_worker.taskError.connect(pnl.on_error)
        self.blast_worker.start()

    def _start_blast_file(self, files: list, cfg: dict):
        pnl = self._panel_blast
        other = getattr(self, "blast_worker", None)
        if other is not None and other.isRunning():
            pnl.on_file_error(
                "The 'BLAST API Search' tab is still running. Wait for it to "
                "finish, or click its Stop button, before parsing a result file.")
            return
        outdir = self._ask_outdir("blastfile")
        if not outdir:
            return
        try:
            os.makedirs(outdir, exist_ok=True)
        except Exception as e:
            pnl.on_file_error(f"Could not create output folder: {e}")
            return
        cfg["outdir"] = outdir
        pnl.set_file_running(True)
        pnl.update_file_status("result", f"Output    │ {outdir}")

        try:
            pnl.stopFileRequested.disconnect()
        except (RuntimeError, TypeError):
            pass
        self._drop_worker("blast_file_worker")

        self.blast_file_worker = _BlastFileWorker(files, cfg)
        pnl.stopFileRequested.connect(self.blast_file_worker.stop)
        self.blast_file_worker.statusUpdated.connect(pnl.update_file_status)
        self.blast_file_worker.progressUpdated.connect(pnl.set_file_progress)
        self.blast_file_worker.taskFinished.connect(pnl.on_file_finished)
        self.blast_file_worker.taskError.connect(pnl.on_file_error)
        self.blast_file_worker.start()

    # ── Shutdown ─────────────────────────────────────────────────────────────

    def closeEvent(self, event):
        # Signal every running worker to stop, then share one short wait. A
        # QThread destroyed while running aborts the process, so one still
        # alive after the wait is handed to C++ ownership and never deleted.
        workers = [getattr(self, a, None) for a in
                   ("blast_worker", "blast_file_worker", "comp_worker")]
        workers += self.__dict__.get("_retired_blast_workers", [])
        for pnl in (self._panel_fasta_tools, self._panel_fastq_inspector,
                    self._panel_bold_formatter, self._panel_genbank):
            workers.append(getattr(pnl, "_worker", None))
            workers += list(getattr(pnl, "_retired_workers", []))
        running = []
        for w in workers:
            try:
                if w is not None and w.isRunning():
                    if hasattr(w, "stop"):
                        w.stop()
                    running.append(w)
            except Exception:
                pass
        deadline = time.monotonic() + 3.0
        for w in running:
            try:
                w.wait(max(0, int((deadline - time.monotonic()) * 1000)))
                if w.isRunning():
                    try:
                        from PyQt5 import sip
                    except ImportError:
                        import sip
                    sip.transferto(w, None)
            except Exception:
                pass

        _settings().setValue("mainwindow/geometry", self.saveGeometry())
        event.accept()


def main():
    multiprocessing.freeze_support()  # required for frozen (PyInstaller) multiprocessing

    # On Windows, give the app its own taskbar identity so the taskbar shows our
    # icon instead of the generic Python one. Must run before any window is shown.
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "BarcodeSuite.app")
        except Exception:
            pass

    app = QtWidgets.QApplication(sys.argv)
    # Theme hooks must be in place before any panel builds its style sheets.
    suite_ui.install_hooks()
    suite_ui.install_compare_results_hook(_compare_panel._CompareResultsWindow)
    apply_theme(app, _settings().value("ui/theme", "light", type=str), STYLESHEET)

    win = MainWindow()
    icon_path = os.path.join(_get_base_dir(), "icon.ico")
    if os.path.exists(icon_path):
        icon = QtGui.QIcon(icon_path)
        app.setWindowIcon(icon)
        win.setWindowIcon(icon)

    geom = _settings().value("mainwindow/geometry")
    if geom:
        win.restoreGeometry(geom)
    else:
        win.resize(1450, 1020)

    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
