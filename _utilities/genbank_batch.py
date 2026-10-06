#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GenBank Batch Search panel (BarcodeSuite-only utility)."""
from __future__ import annotations
import os
import re
import csv
import time
import datetime
import threading
import xlsxwriter
from collections import Counter
from typing import Dict, List, Optional, Tuple
from PyQt5 import QtCore, QtGui, QtWidgets
from .shared import *
from .shared import _get_base_dir, _profiles_dir, _tr, _json_mod


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
