"""
DBF/CSV Segregator — Desktop UI
Double-click app.py to launch. No command line needed.

Requirements:
    pip install PyQt5 openpyxl
"""

import sys
import os
import time
import threading
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QProgressBar, QFrame,
    QScrollArea, QMessageBox, QLineEdit, QGroupBox, QGridLayout
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon

# ── Import core logic from segregate.py ─────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from segregate import load_pan_org_map, process_dbf, process_csv


# ════════════════════════════════════════════════════════════════════════════
#  WORKER THREAD — runs processing in background so UI stays responsive
# ════════════════════════════════════════════════════════════════════════════

class WorkerSignals(QObject):
    progress     = pyqtSignal(int, str)   # (percent, status message)
    finished     = pyqtSignal(dict)       # summary dict
    error        = pyqtSignal(str)        # error message


class ProcessWorker(QThread):
    def __init__(self, master, dbf_files, csv_files, output_dir, wbr22_files=None):
        super().__init__()
        self.master      = master
        self.dbf_files   = dbf_files
        self.csv_files   = csv_files
        self.output_dir  = output_dir
        self.wbr22_files = wbr22_files or set()
        self.signals     = WorkerSignals()

    def run(self):
        try:
            total_files = len(self.dbf_files) + len(self.csv_files)
            t0 = time.perf_counter()

            # Step 1 — Load master
            self.signals.progress.emit(5, "Loading master file...")
            pan_org, folio_product_org = load_pan_org_map(self.master)
            self.signals.progress.emit(20, f"Master loaded — {len(pan_org):,} PAN mappings")

            # Step 2 — Process files
            summaries  = []
            all_tasks  = (
                [(p, "dbf") for p in self.dbf_files] +
                [(p, "csv") for p in self.csv_files]
            )

            for i, (p, ftype) in enumerate(all_tasks):
                fname = Path(p).name
                pct   = 20 + int((i / total_files) * 75)
                self.signals.progress.emit(pct, f"Processing {fname}...")
                if ftype == "dbf":
                    is_wbr22 = p in self.wbr22_files
                    s = process_dbf(p, pan_org, self.output_dir, folio_product_org, wbr22_mode=is_wbr22)
                else:
                    s = process_csv(p, pan_org, self.output_dir, folio_product_org)
                summaries.append(s)

            self.signals.progress.emit(100, "Done!")

            # Compile summary
            elapsed         = time.perf_counter() - t0
            total_records   = sum(s["total"]   for s in summaries)
            total_matched   = sum(s["matched"] for s in summaries)
            total_unmatched = sum(s["unmatched"] for s in summaries)
            total_outputs   = sum(len(s["outputs"]) for s in summaries)
            total_fp        = sum(s.get("fp_matched",  0) for s in summaries)
            total_pan       = sum(s.get("pan_matched", 0) for s in summaries)
            errors          = [s for s in summaries if s["error"]]

            self.signals.finished.emit({
                "total_records"  : total_records,
                "total_matched"  : total_matched,
                "total_unmatched": total_unmatched,
                "total_outputs"  : total_outputs,
                "total_fp"       : total_fp,
                "total_pan"      : total_pan,
                "elapsed"        : elapsed,
                "errors"         : errors,
                "output_dir"     : self.output_dir,
                "per_file"       : summaries,
            })

        except Exception as e:
            self.signals.error.emit(str(e))


# ════════════════════════════════════════════════════════════════════════════
#  FILE ROW WIDGET — shows a selected file with a remove button
# ════════════════════════════════════════════════════════════════════════════

class FileRow(QFrame):
    removed = pyqtSignal(str)

    def __init__(self, filepath: str, show_wbr22: bool = False, parent=None):
        super().__init__(parent)
        self.filepath = filepath
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet("""
            FileRow {
                background: #F8F9FA;
                border: 1px solid #DEE2E6;
                border-radius: 6px;
                padding: 2px;
            }
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 6, 6, 6)

        icon = QLabel("📄")
        icon.setFixedWidth(20)
        layout.addWidget(icon)

        name = QLabel(Path(filepath).name)
        name.setToolTip(filepath)
        name.setFont(QFont("Segoe UI", 10))
        name.setStyleSheet("color: #212529;")
        layout.addWidget(name, 1)

        size_bytes = Path(filepath).stat().st_size
        size_str   = f"{size_bytes/1024/1024:.1f} MB" if size_bytes > 1024*1024 else f"{size_bytes/1024:.1f} KB"
        size_lbl   = QLabel(size_str)
        size_lbl.setStyleSheet("color: #6C757D; font-size: 10px;")
        layout.addWidget(size_lbl)

        # WBR22 checkbox — only shown for DBF files
        self.wbr22_chk = None
        if show_wbr22:
            from PyQt5.QtWidgets import QCheckBox
            self.wbr22_chk = QCheckBox("WBR22")
            self.wbr22_chk.setToolTip("Enable WBR22 mode — FOLIO+PRODUCT only, no PAN fallback")
            self.wbr22_chk.setStyleSheet("""
                QCheckBox {
                    font-size: 10px;
                    color: #6C757D;
                    spacing: 4px;
                }
                QCheckBox::indicator { width: 14px; height: 14px; }
                QCheckBox:checked { color: #0D6EFD; font-weight: bold; }
            """)
            layout.addWidget(self.wbr22_chk)

        btn = QPushButton("✕")
        btn.setFixedSize(24, 24)
        btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                color: #ADB5BD;
                font-size: 14px;
            }
            QPushButton:hover { color: #DC3545; }
        """)
        btn.clicked.connect(lambda: self.removed.emit(self.filepath))
        layout.addWidget(btn)

    def is_wbr22(self) -> bool:
        return self.wbr22_chk is not None and self.wbr22_chk.isChecked()


# ════════════════════════════════════════════════════════════════════════════
#  FILE LIST PANEL — group box with add/clear buttons and file rows
# ════════════════════════════════════════════════════════════════════════════

class FileListPanel(QGroupBox):
    def __init__(self, title: str, ext_filter: str, parent=None):
        super().__init__(title, parent)
        self.ext_filter = ext_filter
        self.files      = []

        self.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.setStyleSheet("""
            QGroupBox {
                border: 1.5px solid #CED4DA;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                color: #495057;
            }
        """)

        outer = QVBoxLayout(self)
        outer.setSpacing(8)

        # Buttons row
        btn_row = QHBoxLayout()
        self.add_btn = QPushButton("+ Add Files")
        self.add_btn.setStyleSheet(self._btn_style("#0D6EFD", "#0B5ED7"))
        self.add_btn.clicked.connect(self._add_files)
        btn_row.addWidget(self.add_btn)

        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.setStyleSheet(self._btn_style("#6C757D", "#5C636A"))
        self.clear_btn.clicked.connect(self._clear_all)
        btn_row.addWidget(self.clear_btn)
        btn_row.addStretch()

        self.count_lbl = QLabel("No files selected")
        self.count_lbl.setStyleSheet("color: #6C757D; font-size: 10px;")
        btn_row.addWidget(self.count_lbl)
        outer.addLayout(btn_row)

        # Scroll area for file rows
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFixedHeight(130)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        self.list_widget = QWidget()
        self.list_layout = QVBoxLayout(self.list_widget)
        self.list_layout.setAlignment(Qt.AlignTop)
        self.list_layout.setSpacing(4)
        self.list_layout.setContentsMargins(0, 0, 0, 0)

        self.empty_lbl = QLabel("No files added yet — click '+ Add Files'")
        self.empty_lbl.setStyleSheet("color: #ADB5BD; font-size: 10px;")
        self.empty_lbl.setAlignment(Qt.AlignCenter)
        self.list_layout.addWidget(self.empty_lbl)

        scroll.setWidget(self.list_widget)
        outer.addWidget(scroll)

    def _btn_style(self, bg, hover):
        return f"""
            QPushButton {{
                background: {bg};
                color: white;
                border: none;
                border-radius: 5px;
                padding: 5px 14px;
                font-size: 10px;
            }}
            QPushButton:hover {{ background: {hover}; }}
            QPushButton:disabled {{ background: #CED4DA; color: #6C757D; }}
        """

    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, f"Select {self.ext_filter.upper()} files", "", f"{self.ext_filter.upper()} Files (*.{self.ext_filter})")
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                row = FileRow(p, show_wbr22=(self.ext_filter == "dbf"))
                row.removed.connect(self._remove_file)
                self.list_layout.addWidget(row)
        self._update_count()

    def _remove_file(self, filepath):
        self.files.remove(filepath)
        for i in range(self.list_layout.count()):
            w = self.list_layout.itemAt(i).widget()
            if isinstance(w, FileRow) and w.filepath == filepath:
                w.deleteLater()
                break
        self._update_count()

    def _clear_all(self):
        self.files.clear()
        for i in reversed(range(self.list_layout.count())):
            w = self.list_layout.itemAt(i).widget()
            if isinstance(w, FileRow):
                w.deleteLater()
        self._update_count()

    def _update_count(self):
        n = len(self.files)
        self.count_lbl.setText(f"{n} file{'s' if n != 1 else ''} selected")
        self.empty_lbl.setVisible(n == 0)

    def get_wbr22_files(self) -> set:
        """Return set of filepaths marked as WBR22."""
        result = set()
        for i in range(self.list_layout.count()):
            w = self.list_layout.itemAt(i).widget()
            if isinstance(w, FileRow) and w.is_wbr22():
                result.add(w.filepath)
        return result


# ════════════════════════════════════════════════════════════════════════════
#  MAIN WINDOW
# ════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DBF / CSV Segregator")
        self.setMinimumSize(700, 720)
        self.worker = None

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(16)

        # ── Header ───────────────────────────────────────────────────────────
        header = QLabel("DBF / CSV Segregator")
        header.setFont(QFont("Segoe UI", 18, QFont.Bold))
        header.setStyleSheet("color: #212529;")
        sub = QLabel("Segregate DBF and CSV files by organisation using a master XLSX mapping")
        sub.setStyleSheet("color: #6C757D; font-size: 11px;")
        main_layout.addWidget(header)
        main_layout.addWidget(sub)

        # ── Master file ───────────────────────────────────────────────────────
        master_group = QGroupBox("Master File (XLSX)")
        master_group.setFont(QFont("Segoe UI", 10, QFont.Bold))
        master_group.setStyleSheet("""
            QGroupBox {
                border: 1.5px solid #CED4DA;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                color: #495057;
            }
        """)
        master_inner = QHBoxLayout(master_group)
        master_inner.setContentsMargins(12, 12, 12, 12)

        self.master_path = QLineEdit()
        self.master_path.setPlaceholderText("No file selected...")
        self.master_path.setReadOnly(True)
        self.master_path.setStyleSheet("""
            QLineEdit {
                border: 1px solid #CED4DA;
                border-radius: 5px;
                padding: 6px 10px;
                font-size: 10px;
                color: #495057;
                background: #F8F9FA;
            }
        """)
        master_inner.addWidget(self.master_path, 1)

        browse_btn = QPushButton("Browse")
        browse_btn.setStyleSheet("""
            QPushButton {
                background: #198754;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 6px 18px;
                font-size: 10px;
            }
            QPushButton:hover { background: #157347; }
        """)
        browse_btn.clicked.connect(self._browse_master)
        master_inner.addWidget(browse_btn)
        main_layout.addWidget(master_group)

        # ── DBF Files ─────────────────────────────────────────────────────────
        self.dbf_panel = FileListPanel("DBF Files", "dbf")
        main_layout.addWidget(self.dbf_panel)

        # ── CSV Files ─────────────────────────────────────────────────────────
        self.csv_panel = FileListPanel("CSV Files", "csv")
        main_layout.addWidget(self.csv_panel)

        # ── Output folder ─────────────────────────────────────────────────────
        out_group = QGroupBox("Output Folder")
        out_group.setFont(QFont("Segoe UI", 10, QFont.Bold))
        out_group.setStyleSheet(master_group.styleSheet())
        out_inner = QHBoxLayout(out_group)
        out_inner.setContentsMargins(12, 12, 12, 12)

        self.output_path = QLineEdit()
        self.output_path.setText(str(Path("./output").resolve()))
        self.output_path.setStyleSheet(self.master_path.styleSheet())
        out_inner.addWidget(self.output_path, 1)

        out_browse = QPushButton("Browse")
        out_browse.setStyleSheet(browse_btn.styleSheet())
        out_browse.clicked.connect(self._browse_output)
        out_inner.addWidget(out_browse)
        main_layout.addWidget(out_group)

        # ── Progress bar ──────────────────────────────────────────────────────
        self.status_lbl = QLabel("Ready")
        self.status_lbl.setStyleSheet("color: #6C757D; font-size: 10px;")
        main_layout.addWidget(self.status_lbl)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFixedHeight(22)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid #CED4DA;
                border-radius: 5px;
                background: #F8F9FA;
                text-align: center;
                font-size: 10px;
                color: #495057;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #0D6EFD, stop:1 #0DCAF0);
                border-radius: 4px;
            }
        """)
        main_layout.addWidget(self.progress)

        # ── Run button ────────────────────────────────────────────────────────
        self.run_btn = QPushButton("▶  Start Segregation")
        self.run_btn.setFixedHeight(44)
        self.run_btn.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.run_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #0D6EFD, stop:1 #0DCAF0);
                color: white;
                border: none;
                border-radius: 8px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #0B5ED7, stop:1 #0BB5D4);
            }
            QPushButton:disabled { background: #CED4DA; color: #6C757D; }
        """)
        self.run_btn.clicked.connect(self._start)
        main_layout.addWidget(self.run_btn)

        main_layout.addStretch()

    def _browse_master(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Master XLSX", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.master_path.setText(path)

    def _browse_output(self):
        path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if path:
            self.output_path.setText(path)

    def _start(self):
        master = self.master_path.text().strip()
        dbf_files = self.dbf_panel.files
        csv_files = self.csv_panel.files
        output_dir = self.output_path.text().strip() or "./output"

        # Validate
        if not master:
            QMessageBox.warning(self, "Missing Input", "Please select a Master XLSX file.")
            return
        if not Path(master).exists():
            QMessageBox.warning(self, "File Not Found", f"Master file not found:\n{master}")
            return
        if not dbf_files and not csv_files:
            QMessageBox.warning(self, "Missing Input", "Please add at least one DBF or CSV file.")
            return

        # Disable UI
        self.run_btn.setEnabled(False)
        self.run_btn.setText("⏳  Processing...")
        self.progress.setValue(0)

        # Start worker
        wbr22_files = self.dbf_panel.get_wbr22_files()
        self.worker = ProcessWorker(master, dbf_files, csv_files, output_dir, wbr22_files)
        self.worker.signals.progress.connect(self._on_progress)
        self.worker.signals.finished.connect(self._on_finished)
        self.worker.signals.error.connect(self._on_error)
        self.worker.start()

    def _on_progress(self, pct, msg):
        self.progress.setValue(pct)
        self.status_lbl.setText(msg)

    def _on_finished(self, s):
        self.progress.setValue(100)
        self.run_btn.setEnabled(True)
        self.run_btn.setText("▶  Start Segregation")

        error_text = ""
        if s["errors"]:
            error_text = "\n\n⚠ Errors:\n" + "\n".join(f"  • {e['file']}: {e['error']}" for e in s["errors"])

        msg = (
            f"✅  Segregation Complete!\n\n"
            f"  Total records    :  {s['total_records']:,}\n"
            f"  Matched rows     :  {s['total_matched']:,}\n"
            f"    via FOLIO+PROD :  {s['total_fp']:,}\n"
            f"    via PAN        :  {s['total_pan']:,}\n"
            f"  Unmatched rows   :  {s['total_unmatched']:,}\n"
            f"  Output files     :  {s['total_outputs']:,}\n"
            f"  Time elapsed     :  {s['elapsed']:.1f}s\n"
            f"\n📁  Output folder:\n  {s['output_dir']}"
            f"{error_text}"
        )

        box = QMessageBox(self)
        box.setWindowTitle("Segregation Complete")
        box.setText(msg)
        box.setStandardButtons(QMessageBox.Ok | QMessageBox.Open)
        box.button(QMessageBox.Open).setText("Open Output Folder")
        result = box.exec_()
        if result == QMessageBox.Open:
            os.startfile(s["output_dir"])

    def _on_error(self, msg):
        self.progress.setValue(0)
        self.run_btn.setEnabled(True)
        self.run_btn.setText("▶  Start Segregation")
        self.status_lbl.setText("Error occurred")
        QMessageBox.critical(self, "Error", f"An error occurred:\n\n{msg}")


# ════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Clean light palette
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#FFFFFF"))
    palette.setColor(QPalette.WindowText, QColor("#212529"))
    palette.setColor(QPalette.Base, QColor("#F8F9FA"))
    palette.setColor(QPalette.AlternateBase, QColor("#E9ECEF"))
    palette.setColor(QPalette.Button, QColor("#F8F9FA"))
    palette.setColor(QPalette.ButtonText, QColor("#212529"))
    app.setPalette(palette)

    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
