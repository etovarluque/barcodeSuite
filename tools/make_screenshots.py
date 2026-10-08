"""Regenerate guide/screenshots/*.png (1440x900, light theme unless noted).

    python tools/make_screenshots.py [--data DIR] [--out DIR] [--only a,b]

--data is ONTbarcoder3's test_files/wcs_l1 (two Conventional runs, a merged
FASTA and a BLAST run). Dialogs that would block are auto-accepted and the
saved theme/sidebar settings are restored at the end.
"""
import argparse
import multiprocessing
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    sys.path.insert(0, ROOT)
    os.chdir(ROOT)
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default=r"E:\Biotools_LGC\ONTbarcoder_v3\test_files\wcs_l1")
    ap.add_argument("--out", default=os.path.join(ROOT, "guide", "screenshots"))
    ap.add_argument("--only", default="")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)
    runs = os.path.join(args.data, "runs")
    run_a = os.path.join(runs, "ont-barcoder_20260930_114431_conv", "consensus_filtered.fa")
    run_b = os.path.join(runs, "ont-barcoder_20260930_114615_conv", "consensus_filtered.fa")
    merge = os.path.join(runs, "ont-barcoder_20260930_121843_merge", "unique_consensus_filtered.fasta")
    blast_fa = os.path.join(runs, "ont-barcoder_20260930_121921_blast", "blast-20260930-121922.fa")
    fastq = os.path.join(args.data, "dataset", "wcs_vert_l1_cytb.fastq")
    tmp = os.path.join(os.path.splitdrive(ROOT)[0] + os.sep, "demo", "out")
    os.makedirs(tmp, exist_ok=True)

    from PyQt5 import QtCore, QtGui, QtWidgets
    import suite_ui
    import barcodeSuite as B

    import traceback
    sys.excepthook = lambda t, v, tb: (traceback.print_exception(t, v, tb, file=sys.stdout), sys.stdout.flush())
    app = QtWidgets.QApplication(sys.argv)
    suite_ui.install_hooks()
    suite_ui.install_compare_results_hook(B._compare_panel._CompareResultsWindow)
    cfg = B._settings()
    saved = {k: cfg.value(k) for k in ("ui/theme", "ui/sidebar_collapsed", "ui/current_tool")}
    suite_ui.apply_theme(app, "light", B.STYLESHEET)
    cfg.setValue("ui/sidebar_collapsed", False)

    QtWidgets.QDialog.exec_ = lambda self: QtWidgets.QDialog.Accepted
    QtWidgets.QMessageBox.exec_ = lambda self: QtWidgets.QMessageBox.Yes
    for n in ("warning", "information", "critical"):
        setattr(QtWidgets.QMessageBox, n, staticmethod(lambda *a, **k: QtWidgets.QMessageBox.Ok))
    QtWidgets.QMessageBox.question = staticmethod(lambda *a, **k: QtWidgets.QMessageBox.Yes)

    w = B.MainWindow()
    w._ask_outdir = lambda suffix: os.path.join(tmp, "out_" + suffix)
    w.setWindowIcon(QtGui.QIcon(os.path.join(ROOT, "icon.ico")))
    w.resize(1440, 900)
    w.show()
    only = [s for s in args.only.split(",") if s]

    def wait(ms):
        loop = QtCore.QEventLoop()
        QtCore.QTimer.singleShot(ms, loop.quit)
        loop.exec_()

    def wait_until(cond, timeout, what):
        t0 = time.time()
        while not cond():
            if time.time() - t0 > timeout:
                raise TimeoutError(what)
            wait(300)

    def save(widget, name, size=(1440, 900)):
        if only and name not in only:
            return
        wait(700)
        pm = widget.grab()
        img = pm.toImage().scaled(size[0], size[1], QtCore.Qt.IgnoreAspectRatio,
                                  QtCore.Qt.SmoothTransformation) \
            if (pm.width(), pm.height()) != tuple(size) else pm.toImage()
        img.save(os.path.join(args.out, name + ".png"))
        print(name, img.width(), img.height())

    def scroll_to(panel, widget):
        area = panel if isinstance(panel, QtWidgets.QScrollArea) else \
            panel.findChild(QtWidgets.QScrollArea)
        if area is not None:
            wait(400)
            y = widget.mapTo(area.widget(), QtCore.QPoint(0, 0)).y()
            area.verticalScrollBar().setValue(max(0, y - 120))

    # 01 FASTQ Inspector
    fq = w._panel_fastq_inspector
    w._switch_tool("fastq")
    fq._on_files([fastq])
    fq._start()
    wait_until(lambda: fq._last_result is not None, 600, "FASTQ Inspector")
    save(w, "01-fastq-inspector")

    # 02 FASTA Tools (Extract -> by header fields)
    ft = w._panel_fasta_tools
    w._switch_tool("fasta")
    ft._drop._add_files([run_a, run_b])
    ft._radio_extract.click()
    ft._radio_filter_fields.click()
    ft._ff_sep._combo.setCurrentIndex(1)  # ";"
    ft._add_filter_criterion()
    r1 = ft._filter_criterion_rows[0]
    r1[1]._mode.setCurrentIndex(1)
    r1[1]._key.setCurrentText("ambs")
    r1[2].setCurrentText("=")
    r1[3].setText("0")
    scroll_to(ft, ft._filter_widget)
    save(w, "02-fasta-tools")

    # 03 FASTA Compare + 04 results window
    cp = w._panel_compare
    w._switch_tool("compare")
    cp._drop._add_files([run_a, run_b])
    cp._id_delim_combo.setCurrentIndex(cp._id_delim_combo.count() - 1)  # Custom...
    cp._id_delim_edit.setText("_all.fa")
    wait(500)
    save(w, "03-fasta-compare")
    if not only or "04-compare-results" in only:
        cp._emit_compare()
        wait_until(lambda: cp._results_win is not None,
                   60, "Compare results")
        cp._open_results_win(); wait(500); rw = cp._results_win
        rw.resize(1300, 760)
        save(rw, "04-compare-results", (1300, 760))
        rw.close()

    # 05 BLAST
    bl = w._panel_blast
    w._switch_tool("blast")
    bl._drop._add_files([blast_fa])
    save(w, "05-blast")

    # 06 GenBank Batch
    w._switch_tool("genbank")
    save(w, "06-genbank-batch")

    # 07 BOLD Formatter (synthetic BOLD-style export)
    import openpyxl
    bold = os.path.join(tmp, "bold_barcode_id.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["BOLD Barcode ID export"])
    ws.append(["Query ID", "ID%", "Phylum", "Class", "Order", "Family", "Genus", "Species"])
    for q in ("S1", "S2"):
        for i in range(6):
            ws.append([q, 100 - i, "Chordata", "Aves", "Passeriformes", "Turdidae", "Turdus", "Turdus sp."])
    wb.save(bold)
    w._switch_tool("bold")
    w._panel_bold_formatter._file_edit.setText(bold)
    save(w, "07-bold-formatter")

    # 08 dark theme, 09 collapsed sidebar
    w._switch_tool("fasta")
    w._toggle_theme()
    save(w, "08-dark-theme")
    w._toggle_theme()
    w._sidebar.set_collapsed(True)
    save(w, "09-collapsed-sidebar")

    for k, v in saved.items():
        if v is not None:
            cfg.setValue(k, v)
    print("done")
    os._exit(0)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
