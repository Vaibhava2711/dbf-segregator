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
import base64
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QProgressBar, QFrame,
    QScrollArea, QMessageBox, QLineEdit, QGroupBox, QGridLayout,
    QDialog, QTableWidget, QTableWidgetItem, QHeaderView,
    QDialogButtonBox, QFormLayout, QCheckBox, QAbstractItemView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon

# ── Import core logic from segregate.py ─────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from segregate import load_pan_org_map, process_dbf, process_csv


# ── Load config.json ─────────────────────────────────────────────────────────
def load_config() -> dict:
    """Load config.json from same directory as app.py."""
    config_path = Path(__file__).parent / "config.json"
    if config_path.exists():
        try:
            import json
            with open(config_path) as f:
                return json.load(f)
        except Exception:
            pass
    return {}

APP_CONFIG = load_config()


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
#  EMAIL FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════

def load_org_email_map(email_excel_path: str) -> dict:
    """
    Load ORG→Email mapping from Excel file.
    Auto-detects columns containing 'ORG' and 'EMAIL'.
    Multiple emails per org supported (comma separated or multiple rows).
    Returns { org_title_case: [email1, email2, ...] }
    """
    import openpyxl
    wb = openpyxl.load_workbook(email_excel_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    org_idx = email_idx = None
    header_row = None
    for rn, row in enumerate(rows[:20]):
        headers = [str(c).strip() if c is not None else "" for c in row]
        oi = next((i for i, h in enumerate(headers) if "ORG" in h.upper()), None)
        ei = next((i for i, h in enumerate(headers) if "EMAIL" in h.upper() or "MAIL" in h.upper()), None)
        if oi is not None and ei is not None:
            org_idx, email_idx, header_row = oi, ei, rn
            break

    if org_idx is None:
        return {}

    org_email = {}
    for row in rows[header_row + 1:]:
        if not any(row):
            continue
        org_val   = row[org_idx]   if org_idx   < len(row) else None
        email_val = row[email_idx] if email_idx < len(row) else None
        if not org_val or not email_val:
            continue
        org_str   = str(org_val).strip()
        email_str = str(email_val).strip()
        if not org_str or not email_str:
            continue
        # Support comma-separated emails in one cell
        emails = [e.strip() for e in email_str.split(",") if e.strip()]
        if org_str not in org_email:
            org_email[org_str] = []
        for e in emails:
            if e not in org_email[org_str]:
                org_email[org_str].append(e)
    return org_email


def find_org_emails(org_name: str, org_email_map: dict) -> list:
    """Case-insensitive lookup of org emails."""
    # Try exact match first
    if org_name in org_email_map:
        return org_email_map[org_name]
    # Try case-insensitive match
    org_upper = org_name.upper()
    for key, emails in org_email_map.items():
        if key.upper() == org_upper:
            return emails
    return []


def send_org_email(smtp_host: str, smtp_port: int, smtp_user: str, smtp_pass: str,
                   from_email: str, to_emails: list,
                   org_name: str, files: list) -> str:
    """
    Send one email to org with all files attached via direct SMTP.
    Returns "" on success or error message string.
    """
    try:
        msg = MIMEMultipart()
        CC_EMAIL = "IFA.ops@scripbox.com"
        msg["From"]    = from_email
        msg["To"]      = ", ".join(to_emails)
        msg["Cc"]      = CC_EMAIL
        from datetime import datetime
        date_str = datetime.now().strftime("%d-%m-%Y")
        msg["Subject"] = f"{org_name} - RTA feeds on {date_str}"

        body = (
            f"Hi Team,\n\n"
            f"Please find attached the RTA-wise feed file for your reference.\n\n"
            f"Cams: DBF file format.\n"
            f"Karvy: CSV file format.\n\n"
            f"Kindly confirm once all the Feeds are successfully uploaded to your systems.\n\n"
            f"--\n"
            f"Thanks & Regards,\n"
            f"Investment Operations Team\n"
            f"Scripbox"
        )
        msg.attach(MIMEText(body, "plain"))

        for fpath in files:
            with open(fpath, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={Path(fpath).name}")
            msg.attach(part)

        # Try SSL first (port 465), then TLS (port 587)
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode    = ssl.CERT_NONE   # handles self-signed company certs

        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_host, smtp_port, context=ctx) as server:
                server.login(smtp_user, smtp_pass)
                server.sendmail(from_email, to_emails + [CC_EMAIL], msg.as_string())
        else:
            with smtplib.SMTP(smtp_host, smtp_port) as server:
                server.ehlo()
                server.starttls(context=ctx)
                server.login(smtp_user, smtp_pass)
                server.sendmail(from_email, to_emails + [CC_EMAIL], msg.as_string())
        return ""
    except Exception as e:
        return str(e)


# ════════════════════════════════════════════════════════════════════════════
#  EMAIL PREVIEW DIALOG
# ════════════════════════════════════════════════════════════════════════════

class EmailPreviewDialog(QDialog):
    def __init__(self, org_files: dict, org_email_map: dict, output_dir: str, parent=None):
        """
        org_files: { org_name: [filepath, ...] }
        org_email_map: { org_name: [email, ...] }
        """
        super().__init__(parent)
        self.org_files    = org_files
        self.org_email_map = org_email_map
        self.output_dir   = output_dir
        self.setWindowTitle("Email Preview — Confirm Before Sending")
        self.setMinimumSize(800, 520)
        self.setSizeGripEnabled(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        # Header
        hdr = QLabel("Review and confirm emails before sending")
        hdr.setFont(QFont("Segoe UI", 13, QFont.Bold))
        hdr.setStyleSheet("color: #212529;")
        layout.addWidget(hdr)

        sub = QLabel("Uncheck any org you don't want to email. Edit email addresses if needed.")
        sub.setStyleSheet("color: #6C757D; font-size: 11px;")
        layout.addWidget(sub)

        # Gmail credentials
        cred_group = QGroupBox("Gmail Credentials")
        cred_group.setFont(QFont("Segoe UI", 10, QFont.Bold))
        cred_group.setStyleSheet("""
            QGroupBox {
                border: 1.5px solid #CED4DA; border-radius: 8px;
                margin-top: 12px; padding-top: 8px;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 12px; color: #495057; }
        """)
        cred_layout = QFormLayout(cred_group)
        cred_layout.setContentsMargins(12, 12, 12, 12)

        has_config = bool(APP_CONFIG.get("smtp_host"))
        readonly_style = self._input_style() + "QLineEdit { background: #E9ECEF; }"

        self.from_email = QLineEdit()
        self.from_email.setPlaceholderText("vaibhav.goyal@scripbox.com")
        self.from_email.setStyleSheet(self._input_style())
        self.from_email.setText(APP_CONFIG.get("from_email", ""))
        if has_config: self.from_email.setReadOnly(True); self.from_email.setStyleSheet(readonly_style)
        cred_layout.addRow("From email:", self.from_email)

        self.smtp_host = QLineEdit()
        self.smtp_host.setPlaceholderText("mail.scripbox.com")
        self.smtp_host.setStyleSheet(self._input_style())
        self.smtp_host.setText(APP_CONFIG.get("smtp_host", ""))
        if has_config: self.smtp_host.setReadOnly(True); self.smtp_host.setStyleSheet(readonly_style)
        cred_layout.addRow("SMTP host:", self.smtp_host)

        self.smtp_port = QLineEdit()
        self.smtp_port.setPlaceholderText("587")
        self.smtp_port.setStyleSheet(self._input_style())
        self.smtp_port.setText(str(APP_CONFIG.get("smtp_port", "587")))
        if has_config: self.smtp_port.setReadOnly(True); self.smtp_port.setStyleSheet(readonly_style)
        cred_layout.addRow("SMTP port:", self.smtp_port)

        self.smtp_user = QLineEdit()
        self.smtp_user.setPlaceholderText("vaibhav.goyal@scripbox.com")
        self.smtp_user.setStyleSheet(self._input_style())
        self.smtp_user.setText(APP_CONFIG.get("smtp_user", ""))
        if has_config: self.smtp_user.setReadOnly(True); self.smtp_user.setStyleSheet(readonly_style)
        cred_layout.addRow("SMTP username:", self.smtp_user)

        self.smtp_pass = QLineEdit()
        self.smtp_pass.setPlaceholderText("Your email password")
        self.smtp_pass.setEchoMode(QLineEdit.Password)
        self.smtp_pass.setStyleSheet(self._input_style())
        self.smtp_pass.setText(APP_CONFIG.get("smtp_pass", ""))
        if has_config: self.smtp_pass.setReadOnly(True); self.smtp_pass.setStyleSheet(readonly_style)
        cred_layout.addRow("SMTP password:", self.smtp_pass)

        hint = QLabel("ℹ  Credentials loaded from config.json" if has_config else "ℹ  Ask your IT team for SMTP host, port, username and password")
        hint.setStyleSheet("color: #198754; font-size: 10px;" if has_config else "color: #6C757D; font-size: 10px;")
        cred_layout.addRow("", hint)
        layout.addWidget(cred_group)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Send", "Organisation", "Email(s)", "Files"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.setSelectionMode(QAbstractItemView.NoSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget { border: 1px solid #CED4DA; border-radius: 6px; font-size: 11px; }
            QHeaderView::section { background: #F8F9FA; font-weight: bold; padding: 6px; border: none; border-bottom: 1px solid #CED4DA; }
        """)

        all_orgs = sorted(org_files.keys())
        self.table.setRowCount(len(all_orgs))
        self.row_checks = {}

        for r, org in enumerate(all_orgs):
            # Checkbox
            chk = QCheckBox()
            emails = find_org_emails(org, org_email_map)
            chk.setChecked(len(emails) > 0)
            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.addWidget(chk)
            chk_layout.setAlignment(Qt.AlignCenter)
            chk_layout.setContentsMargins(0,0,0,0)
            self.table.setCellWidget(r, 0, chk_widget)
            self.row_checks[org] = chk

            # Org name
            org_item = QTableWidgetItem(org)
            org_item.setFlags(org_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(r, 1, org_item)

            # Emails (editable)
            email_item = QTableWidgetItem(", ".join(emails) if emails else "⚠ No email found")
            if not emails:
                email_item.setForeground(QColor("#DC3545"))
            self.table.setItem(r, 2, email_item)

            # File count
            fcount = len(org_files.get(org, []))
            fc_item = QTableWidgetItem(f"{fcount} file{'s' if fcount != 1 else ''}")
            fc_item.setFlags(fc_item.flags() & ~Qt.ItemIsEditable)
            fc_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, 3, fc_item)

        self.all_orgs = all_orgs
        layout.addWidget(self.table)

        # Status label
        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("font-size: 11px; color: #495057;")
        layout.addWidget(self.status_lbl)

        # Progress bar
        self.send_progress = QProgressBar()
        self.send_progress.setValue(0)
        self.send_progress.setFixedHeight(18)
        self.send_progress.setVisible(False)
        self.send_progress.setStyleSheet("""
            QProgressBar { border: 1px solid #CED4DA; border-radius: 4px; background: #F8F9FA; text-align: center; font-size: 10px; }
            QProgressBar::chunk { background: #198754; border-radius: 3px; }
        """)
        layout.addWidget(self.send_progress)

        # Buttons
        btn_row = QHBoxLayout()
        self.send_btn = QPushButton("📧  Send Emails")
        self.send_btn.setFixedHeight(38)
        self.send_btn.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self.send_btn.setStyleSheet("""
            QPushButton { background: #198754; color: white; border: none; border-radius: 6px; padding: 0 24px; }
            QPushButton:hover { background: #157347; }
            QPushButton:disabled { background: #CED4DA; color: #6C757D; }
        """)
        self.send_btn.clicked.connect(self._send_emails)
        btn_row.addWidget(self.send_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedHeight(38)
        cancel_btn.setStyleSheet("""
            QPushButton { background: #6C757D; color: white; border: none; border-radius: 6px; padding: 0 20px; }
            QPushButton:hover { background: #5C636A; }
        """)
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

    def _input_style(self):
        return """
            QLineEdit {
                border: 1px solid #CED4DA; border-radius: 5px;
                padding: 6px 10px; font-size: 11px; color: #495057; background: #F8F9FA;
            }
        """

    def _send_emails(self):
        from_email = self.from_email.text().strip()
        smtp_host  = self.smtp_host.text().strip()
        smtp_port  = int(self.smtp_port.text().strip() or "587")
        smtp_user  = self.smtp_user.text().strip()
        smtp_pass  = self.smtp_pass.text().strip()

        if not from_email or not smtp_host or not smtp_user or not smtp_pass:
            QMessageBox.warning(self, "Missing Credentials", "Please fill in all SMTP fields.")
            return

        # Collect selected orgs
        to_send = []
        for r, org in enumerate(self.all_orgs):
            chk = self.row_checks[org]
            if not chk.isChecked():
                continue
            email_item = self.table.item(r, 2)
            emails = [e.strip() for e in email_item.text().split(",") if e.strip() and "@" in e]
            if not emails:
                continue
            to_send.append((org, emails, self.org_files.get(org, [])))

        if not to_send:
            QMessageBox.warning(self, "Nothing to Send", "No orgs selected or no valid email addresses found.")
            return

        self.send_btn.setEnabled(False)
        self.send_progress.setVisible(True)
        self.send_progress.setMaximum(len(to_send))
        self.send_progress.setValue(0)

        # Group orgs by email set — same email(s) = one combined email
        from collections import defaultdict
        email_groups = defaultdict(lambda: {"orgs": [], "files": [], "emails": []})
        for org, emails, files in to_send:
            key = tuple(sorted(e.lower() for e in emails))
            email_groups[key]["orgs"].append(org)
            email_groups[key]["files"].extend(files)
            email_groups[key]["emails"] = emails

        self.send_progress.setMaximum(len(email_groups))
        errors = []
        for i, (key, group) in enumerate(email_groups.items()):
            org_label = " & ".join(group["orgs"])
            self.status_lbl.setText(f"Sending to {org_label}...")
            QApplication.processEvents()
            err = send_org_email(smtp_host, smtp_port, smtp_user, smtp_pass, from_email, group["emails"], org_label, group["files"])
            if err:
                errors.append(f"{org_label}: {err}")
            self.send_progress.setValue(i + 1)
            QApplication.processEvents()

        self.send_btn.setEnabled(True)
        self.status_lbl.setText("")

        if errors:
            err_msg = f"{len(to_send) - len(errors)} sent successfully.\n\nErrors:\n" + "\n".join(errors)
            QMessageBox.warning(self, "Some Emails Failed", err_msg)
        else:
            QMessageBox.information(self, "All Emails Sent",
                f"✅ Successfully sent emails to {len(to_send)} organisation(s)!")
            self.accept()


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

        # ── Email config ──────────────────────────────────────────────────────
        email_group = QGroupBox("Email Config (Optional)")
        email_group.setFont(QFont("Segoe UI", 10, QFont.Bold))
        email_group.setStyleSheet(master_group.styleSheet())
        email_inner = QHBoxLayout(email_group)
        email_inner.setContentsMargins(12, 12, 12, 12)

        self.email_map_path = QLineEdit()
        self.email_map_path.setPlaceholderText("Select ORG→Email mapping Excel file (optional)...")
        self.email_map_path.setReadOnly(True)
        self.email_map_path.setStyleSheet(self.master_path.styleSheet())
        email_inner.addWidget(self.email_map_path, 1)

        email_browse = QPushButton("Browse")
        email_browse.setStyleSheet(browse_btn.styleSheet())
        email_browse.clicked.connect(self._browse_email_map)
        email_inner.addWidget(email_browse)
        main_layout.addWidget(email_group)

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

    def _browse_email_map(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Email Mapping Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.email_map_path.setText(path)

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
        self._last_summary = s  # store for email use

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

        # Add Send Emails button if email config is provided
        send_btn = None
        if self.email_map_path.text().strip():
            send_btn = box.addButton("📧 Send Emails", QMessageBox.ActionRole)

        result = box.exec_()
        clicked = box.clickedButton()

        if clicked == box.button(QMessageBox.Open):
            os.startfile(s["output_dir"])
        elif send_btn and clicked == send_btn:
            self._show_email_preview(s)

    def _show_email_preview(self, s):
        """Build org→files map and show email preview dialog."""
        email_excel = self.email_map_path.text().strip()
        if not Path(email_excel).exists():
            QMessageBox.warning(self, "File Not Found", f"Email mapping file not found:\n{email_excel}")
            return

        # Load org→email map
        try:
            org_email_map = load_org_email_map(email_excel)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load email mapping:\n{e}")
            return

        # Build org→files map from per_file summaries
        # Use actual org names from processing summaries (not filename parsing)
        output_dir = Path(s["output_dir"])
        org_files  = {}
        for per_file in s["per_file"]:
            base_name = Path(per_file["file"]).stem
            for fpath in per_file.get("outputs", []):
                p = Path(fpath)
                if not p.exists() or "UNMATCHED" in p.name:
                    continue
                # Extract org name by removing base_name prefix from stem
                stem = p.stem
                # Remove base_name_ prefix to get org name
                prefix = base_name + "_"
                if stem.startswith(prefix):
                    org_raw = stem[len(prefix):]
                else:
                    # fallback: everything after first underscore
                    parts = stem.split("_", 1)
                    org_raw = parts[1] if len(parts) > 1 else stem
                # Convert underscores back to spaces and title case
                org_name = org_raw.replace("_", " ").title()
                if org_name not in org_files:
                    org_files[org_name] = []
                org_files[org_name].append(fpath)

        if not org_files:
            QMessageBox.warning(self, "No Files", "No output files found to send.")
            return

        dlg = EmailPreviewDialog(org_files, org_email_map, s["output_dir"], self)
        dlg.exec_()

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
