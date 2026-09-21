import sys
import os
import csv
import shutil
import sqlite3
import zipfile
import logging
import json
from datetime import datetime, date
from pathlib import Path

from PySide6.QtCore import Qt, QRectF
from PySide6.QtGui import QAction, QPixmap, QPainter, QPen, QBrush, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFormLayout, QLineEdit, QTextEdit, QPushButton, QLabel,
    QComboBox, QTableWidget, QTableWidgetItem, QMessageBox,
    QFileDialog, QDateEdit, QSpinBox, QDoubleSpinBox, QGroupBox,
    QSplitter, QListWidget, QStackedWidget, QDialog, QDialogButtonBox,
    QHeaderView, QAbstractItemView, QCheckBox, QScrollArea, QSizePolicy
)

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import Workbook


APP_NAME = "HSE Management System"
APP_VERSION = "1.1.0"

APP_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / APP_NAME
DB_DIR = APP_DIR / "database"
ATTACH_DIR = APP_DIR / "attachments"
REPORT_DIR = APP_DIR / "reports"
BACKUP_DIR = APP_DIR / "backups"
LOG_DIR = APP_DIR / "logs"

for p in [APP_DIR, DB_DIR, ATTACH_DIR, REPORT_DIR, BACKUP_DIR, LOG_DIR]:
    p.mkdir(parents=True, exist_ok=True)

DB_FILE = DB_DIR / "hse.db"

logging.basicConfig(
    filename=LOG_DIR / "application.log",
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)


# ============================================================
# DATABASE
# ============================================================

class Database:

    def __init__(self):
        self.conn = sqlite3.connect(DB_FILE)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.create_tables()

    def create_tables(self):

        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );

        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT,
            name TEXT NOT NULL,
            designation TEXT,
            company TEXT,
            department TEXT,
            contact TEXT,
            active INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            company_type TEXT,
            contact_person TEXT,
            contact_number TEXT,
            status TEXT DEFAULT 'Active'
        );

        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project TEXT NOT NULL,
            area TEXT,
            location TEXT,
            description TEXT,
            active INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS observations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT UNIQUE,
            obs_date TEXT,
            obs_time TEXT,
            project TEXT,
            company TEXT,
            location TEXT,
            area TEXT,
            responsible TEXT,
            designation TEXT,
            employee_id TEXT,
            observer TEXT,
            observer_designation TEXT,
            observer_id TEXT,
            observer_company TEXT,
            obs_type TEXT,
            category TEXT,
            subcategory TEXT,
            observation TEXT,
            immediate_action TEXT,
            corrective_action TEXT,
            preventive_action TEXT,
            priority TEXT,
            target_date TEXT,
            status TEXT DEFAULT 'Open',
            closeout_date TEXT,
            closed_by TEXT,
            verified_by TEXT,
            verification_date TEXT,
            closeout_comments TEXT,
            verification TEXT,
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS observation_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            observation_id INTEGER,
            file_path TEXT,
            attachment_type TEXT,
            FOREIGN KEY(observation_id)
                REFERENCES observations(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS incidents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT UNIQUE,
            incident_date TEXT,
            incident_time TEXT,
            location TEXT,
            project TEXT,
            company TEXT,
            department TEXT,
            activity TEXT,
            incident_type TEXT,
            person_involved TEXT,
            employee_id TEXT,
            designation TEXT,
            supervisor TEXT,
            witnesses TEXT,
            description TEXT,
            immediate_action TEXT,
            consequences TEXT,
            potential_consequences TEXT,
            equipment TEXT,
            investigation_method TEXT,
            why1 TEXT,
            why2 TEXT,
            why3 TEXT,
            why4 TEXT,
            why5 TEXT,
            direct_cause TEXT,
            contributing_factors TEXT,
            root_cause TEXT,
            corrective_action TEXT,
            preventive_action TEXT,
            status TEXT DEFAULT 'Open',
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS audits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT UNIQUE,
            audit_date TEXT,
            audit_type TEXT,
            standard TEXT,
            project TEXT,
            location TEXT,
            department TEXT,
            auditor TEXT,
            lead_auditor TEXT,
            auditee TEXT,
            scope TEXT,
            objective TEXT,
            criteria TEXT,
            start_time TEXT,
            end_time TEXT,
            status TEXT DEFAULT 'Open',
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS audit_findings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            audit_id INTEGER,
            clause TEXT,
            sub_clause TEXT,
            requirement TEXT,
            finding_type TEXT,
            observation TEXT,
            evidence TEXT,
            risk_impact TEXT,
            corrective_action TEXT,
            responsible TEXT,
            target_date TEXT,
            status TEXT DEFAULT 'Open',
            verification TEXT,
            closeout_evidence TEXT,
            FOREIGN KEY(audit_id)
                REFERENCES audits(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS capa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT UNIQUE,
            source TEXT,
            reference_number TEXT,
            finding TEXT,
            root_cause TEXT,
            corrective_action TEXT,
            preventive_action TEXT,
            responsible TEXT,
            priority TEXT,
            target_date TEXT,
            status TEXT DEFAULT 'Open',
            verification TEXT,
            closeout_evidence TEXT,
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT,
            subcategory TEXT,
            active INTEGER DEFAULT 1
        );
        """)

        # Backward-compatible additions for existing installations.
        existing = {r["name"] for r in self.conn.execute("PRAGMA table_info(incidents)").fetchall()}
        if "investigation_details" not in existing:
            self.conn.execute("ALTER TABLE incidents ADD COLUMN investigation_details TEXT")

        self.conn.commit()

    def execute(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        self.conn.commit()
        return cur

    def fetchall(self, sql, params=()):
        return self.conn.execute(sql, params).fetchall()

    def fetchone(self, sql, params=()):
        return self.conn.execute(sql, params).fetchone()

    def setting(self, key, default=""):
        row = self.fetchone(
            "SELECT value FROM settings WHERE key=?",
            (key,)
        )
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.execute(
            """
            INSERT INTO settings(key,value)
            VALUES(?,?)
            ON CONFLICT(key)
            DO UPDATE SET value=excluded.value
            """,
            (key, value)
        )


db = Database()


# ============================================================
# HELPERS
# ============================================================

OBS_TYPES = [
    "Unsafe Act",
    "Unsafe Condition",
    "Good Observation",
    "Good Practice",
    "Positive Observation",
    "Environmental Observation",
    "Near Miss",
    "Other"
]

CATEGORIES = [
    "PPE",
    "Work at Height",
    "Excavation",
    "Electrical Safety",
    "Lifting Operations",
    "Scaffolding",
    "Confined Space",
    "Fire Safety",
    "Housekeeping",
    "Slip Trip Fall",
    "Chemical Safety",
    "Environmental",
    "Traffic Management",
    "Plant and Machinery",
    "Manual Handling",
    "Emergency Preparedness",
    "Heat Stress",
    "Welfare",
    "Dropped Objects",
    "Tools and Equipment",
    "Permit to Work",
    "Barricading",
    "Material Storage",
    "Access and Egress",
    "Gas Testing",
    "Hot Work",
    "Noise",
    "Dust",
    "Waste Management",
    "Environmental Spill",
    "Other"
]

PRIORITIES = ["Low", "Medium", "High", "Critical"]

STATUSES = [
    "Open",
    "In Progress",
    "Pending Verification",
    "Closed",
    "Cancelled"
]

INCIDENT_TYPES = [
    "Fatality",
    "Lost Time Injury",
    "Medical Treatment Case",
    "Restricted Work Case",
    "First Aid Case",
    "Near Miss",
    "Property Damage",
    "Environmental Incident",
    "Fire Incident",
    "Vehicle Incident",
    "Equipment Damage",
    "Chemical Spill",
    "Other"
]

INVESTIGATION_METHODS = [
    "Root Cause Analysis",
    "5 Why Analysis",
    "Fishbone / Ishikawa",
    "ICAM",
    "Barrier Analysis",
    "Bow-Tie Analysis",
    "Fault Tree Analysis",
    "Causal Tree",
    "Other"
]

AUDIT_TYPES = [
    "Internal Audit",
    "External Audit",
    "Client Audit",
    "Certification Audit",
    "Surveillance Audit",
    "Regulatory Audit",
    "Project Audit",
    "Supplier Audit",
    "Other"
]

FINDING_TYPES = [
    "Positive Observation",
    "Good Practice",
    "Conformity",
    "Opportunity for Improvement",
    "Observation",
    "Minor Nonconformity",
    "Major Nonconformity",
    "Environmental Finding",
    "Legal / Compliance Finding",
    "Other"
]

CAPA_SOURCES = [
    "HSE Inspection",
    "Incident",
    "Near Miss",
    "Audit",
    "Client Observation",
    "Regulatory Inspection",
    "Environmental Incident",
    "Management Review",
    "Employee Observation",
    "Other"
]


ISO_45001_CLAUSES = {
    "4 - Context of the organization": [
        ("4.1", "Understanding the organization and its context"),
        ("4.2", "Understanding the needs and expectations of workers and other interested parties"),
        ("4.3", "Determining the scope of the OH&S management system"),
        ("4.4", "OH&S management system"),
    ],
    "5 - Leadership and worker participation": [
        ("5.1", "Leadership and commitment"),
        ("5.2", "OH&S policy"),
        ("5.3", "Organizational roles, responsibilities and authorities"),
        ("5.4", "Consultation and participation of workers"),
    ],
    "6 - Planning": [
        ("6.1.1", "General"),
        ("6.1.2", "Hazard identification and assessment of risks and opportunities"),
        ("6.1.3", "Determination of legal requirements and other requirements"),
        ("6.1.4", "Planning action"),
        ("6.2.1", "OH&S objectives"),
        ("6.2.2", "Planning to achieve OH&S objectives"),
    ],
    "7 - Support": [
        ("7.1", "Resources"), ("7.2", "Competence"), ("7.3", "Awareness"),
        ("7.4.1", "Communication - General"), ("7.4.2", "Internal communication"),
        ("7.4.3", "External communication"), ("7.5.1", "Documented information - General"),
        ("7.5.2", "Creating and updating"), ("7.5.3", "Control of documented information"),
    ],
    "8 - Operation": [
        ("8.1.1", "General"), ("8.1.2", "Eliminating hazards and reducing OH&S risks"),
        ("8.1.3", "Management of change"), ("8.1.4.1", "Procurement - General"),
        ("8.1.4.2", "Contractors"), ("8.1.4.3", "Outsourcing"),
        ("8.2", "Emergency preparedness and response"),
    ],
    "9 - Performance evaluation": [
        ("9.1.1", "Monitoring, measurement, analysis and evaluation - General"),
        ("9.1.2", "Evaluation of compliance"), ("9.2.1", "Internal audit - General"),
        ("9.2.2", "Internal audit programme"), ("9.3", "Management review"),
    ],
    "10 - Improvement": [
        ("10.1", "General"), ("10.2", "Incident, nonconformity and corrective action"),
        ("10.3", "Continual improvement"),
    ],
}

ISO_14001_CLAUSES = {
    "4 - Context of the organization": [
        ("4.1", "Understanding the organization and its context"),
        ("4.2", "Understanding the needs and expectations of interested parties"),
        ("4.3", "Determining the scope of the environmental management system"),
        ("4.4", "Environmental management system"),
    ],
    "5 - Leadership": [
        ("5.1", "Leadership and commitment"), ("5.2", "Environmental policy"),
        ("5.3", "Organizational roles, responsibilities and authorities"),
    ],
    "6 - Planning": [
        ("6.1.1", "General"), ("6.1.2", "Environmental aspects"),
        ("6.1.3", "Compliance obligations"), ("6.1.4", "Planning action"),
        ("6.2.1", "Environmental objectives"), ("6.2.2", "Planning actions to achieve environmental objectives"),
    ],
    "7 - Support": [
        ("7.1", "Resources"), ("7.2", "Competence"), ("7.3", "Awareness"),
        ("7.4.1", "Communication - General"), ("7.4.2", "Internal communication"),
        ("7.4.3", "External communication"), ("7.5.1", "Documented information - General"),
        ("7.5.2", "Creating and updating"), ("7.5.3", "Control of documented information"),
    ],
    "8 - Operation": [
        ("8.1", "Operational planning and control"), ("8.2", "Emergency preparedness and response"),
    ],
    "9 - Performance evaluation": [
        ("9.1.1", "Monitoring, measurement, analysis and evaluation - General"),
        ("9.1.2", "Evaluation of compliance"), ("9.2.1", "Internal audit - General"),
        ("9.2.2", "Internal audit programme"), ("9.3", "Management review"),
    ],
    "10 - Improvement": [
        ("10.1", "General"), ("10.2", "Nonconformity and corrective action"),
        ("10.3", "Continual improvement"),
    ],
}

AUDIT_FINDING_TYPES = [
    "Positive Observation", "Good Practice", "Conformity",
    "Opportunity for Improvement", "Observation", "Minor Nonconformity",
    "Major Nonconformity", "Environmental Finding", "Legal / Compliance Finding", "Other"
]


def next_number(prefix, table):
    year = datetime.now().year
    row = db.fetchone(
        f"""
        SELECT number FROM {table}
        WHERE number LIKE ?
        ORDER BY id DESC LIMIT 1
        """,
        (f"{prefix}-{year}-%",)
    )

    if not row:
        n = 1
    else:
        try:
            n = int(row["number"].split("-")[-1]) + 1
        except Exception:
            n = 1

    return f"{prefix}-{year}-{n:05d}"


def today():
    return datetime.now().strftime("%Y-%m-%d")


def now_time():
    return datetime.now().strftime("%H:%M")


def safe(value):
    return "" if value is None else str(value)


def overdue(target, status):
    if not target or status in ("Closed", "Cancelled"):
        return False
    try:
        return date.fromisoformat(target) < date.today()
    except Exception:
        return False


# ============================================================
# SIMPLE BUILT-IN CHART WIDGETS
# ============================================================

class PieChartWidget(QWidget):
    def __init__(self, values, title="Observation Distribution", parent=None):
        super().__init__(parent)
        self.values = values
        self.title = title
        self.setMinimumHeight(300)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QBrush(Qt.GlobalColor.white))
        painter.setPen(QPen(Qt.GlobalColor.black))
        painter.setFont(QFont("Arial", 12, QFont.Bold))
        painter.drawText(15, 25, self.title)
        total = sum(max(0, v) for v in self.values.values())
        if total <= 0:
            painter.setFont(QFont("Arial", 10))
            painter.drawText(15, 60, "No observation data available")
            painter.end(); return
        rect = QRectF(30, 55, min(self.width()*0.55, 280), min(self.height()-85, 220))
        start = 0.0
        palette = [
            Qt.GlobalColor.darkBlue, Qt.GlobalColor.darkGreen, Qt.GlobalColor.darkRed,
            Qt.GlobalColor.darkCyan, Qt.GlobalColor.darkMagenta, Qt.GlobalColor.darkYellow,
            Qt.GlobalColor.gray
        ]
        for i,(label,value) in enumerate(self.values.items()):
            span = 360.0 * value / total
            painter.setBrush(QBrush(palette[i % len(palette)]))
            painter.drawPie(rect, int(start*16), int(span*16))
            start += span
        x = int(rect.right()+25); y = 70
        painter.setFont(QFont("Arial", 9))
        for i,(label,value) in enumerate(self.values.items()):
            painter.setBrush(QBrush(palette[i % len(palette)])); painter.drawRect(x, y-10, 12, 12)
            painter.setPen(QPen(Qt.GlobalColor.black)); painter.drawText(x+18, y, f"{label}: {value}")
            y += 24
        painter.end()

class BarChartWidget(QWidget):
    def __init__(self, values, title="Observations by Category", parent=None):
        super().__init__(parent)
        self.values = values
        self.title = title
        self.setMinimumHeight(320)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QBrush(Qt.GlobalColor.white))
        painter.setPen(QPen(Qt.GlobalColor.black)); painter.setFont(QFont("Arial", 12, QFont.Bold))
        painter.drawText(15, 25, self.title)
        items = list(self.values.items())
        if not items:
            painter.setFont(QFont("Arial", 10)); painter.drawText(15, 60, "No observation data available"); painter.end(); return
        max_value = max([v for _,v in items] or [1])
        left, top, bottom = 45, 55, self.height()-45
        width = max(1, self.width()-left-20); height = max(1, bottom-top)
        bar_w = max(12, width / max(1,len(items)) - 8)
        painter.setFont(QFont("Arial", 8))
        for i,(label,value) in enumerate(items):
            x = left + i*(width/len(items)) + 4
            h = (value/max_value)*(height-35)
            y = bottom-h
            painter.setBrush(QBrush(Qt.GlobalColor.darkBlue)); painter.drawRect(QRectF(x,y,bar_w,h))
            painter.setPen(QPen(Qt.GlobalColor.black)); painter.drawText(int(x), int(y-5), str(value))
            painter.save(); painter.translate(x+bar_w/2, bottom+8); painter.rotate(-45); painter.drawText(0,0,label[:18]); painter.restore()
        painter.end()

# ============================================================
# MAIN WINDOW
# ============================================================

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle(
            f"{APP_NAME} - {APP_VERSION}"
        )

        self.resize(1400, 850)

        self.setStyleSheet("""
        QMainWindow {
            background: #F4F6F8;
        }

        QLabel {
            color: #202020;
        }

        QPushButton {
            background: #17365D;
            color: white;
            border: none;
            padding: 9px 14px;
            border-radius: 5px;
            min-height: 20px;
        }

        QPushButton:hover {
            background: #245486;
        }

        QLineEdit, QTextEdit, QComboBox, QDateEdit {
            background: white;
            border: 1px solid #B8C2CC;
            border-radius: 4px;
            padding: 7px;
        }

        QTableWidget {
            background: white;
            gridline-color: #D5DADF;
        }

        QHeaderView::section {
            background: #17365D;
            color: white;
            padding: 7px;
            border: none;
        }

        QListWidget {
            background: #102A43;
            color: white;
            border: none;
        }

        QListWidget::item {
            padding: 14px;
        }

        QListWidget::item:selected {
            background: #245486;
        }

        QGroupBox {
            font-weight: bold;
            border: 1px solid #CCD3DA;
            margin-top: 10px;
            padding: 10px;
        }
        """)

        self.build_ui()
        self.dashboard()

    # --------------------------------------------------------

    def build_ui(self):

        central = QWidget()
        self.setCentralWidget(central)

        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)

        self.nav = QListWidget()

        modules = [
            "Dashboard",
            "HSE Inspection Register",
            "Incident Investigation",
            "Audit Register",
            "CAPA Register",
            "Reports",
            "Master Data",
            "Settings"
        ]

        self.nav.addItems(modules)
        self.nav.setFixedWidth(245)
        self.nav.currentRowChanged.connect(
            self.navigation_changed
        )

        layout.addWidget(self.nav)

        self.stack = QStackedWidget()
        layout.addWidget(self.stack, 1)

    # --------------------------------------------------------

    def navigation_changed(self, row):

        if row == 0:
            self.dashboard()
        elif row == 1:
            self.observations()
        elif row == 2:
            self.incidents()
        elif row == 3:
            self.audits()
        elif row == 4:
            self.capa()
        elif row == 5:
            self.reports()
        elif row == 6:
            self.master_data()
        elif row == 7:
            self.settings_page()

    # --------------------------------------------------------

    def page(self, title):

        w = QWidget()
        layout = QVBoxLayout(w)

        title_label = QLabel(title)
        title_label.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #17365D;
            padding: 8px;
        """)

        layout.addWidget(title_label)

        self.stack.addWidget(w)
        self.stack.setCurrentWidget(w)

        return w, layout

    # ========================================================
    # DASHBOARD
    # ========================================================

    def dashboard(self):

        w, layout = self.page("HSE Dashboard")

        cards = QHBoxLayout()
        values = [
            ("Observations", "SELECT COUNT(*) c FROM observations"),
            ("Open Observations", "SELECT COUNT(*) c FROM observations WHERE status NOT IN ('Closed','Cancelled')"),
            ("Overdue", "SELECT COUNT(*) c FROM observations WHERE status NOT IN ('Closed','Cancelled') AND target_date < date('now')"),
            ("Incidents", "SELECT COUNT(*) c FROM incidents"),
            ("Audits", "SELECT COUNT(*) c FROM audits"),
            ("CAPA", "SELECT COUNT(*) c FROM capa"),
        ]
        for name, sql in values:
            value=db.fetchone(sql)["c"]
            box=QGroupBox(name); bl=QVBoxLayout(box)
            label=QLabel(str(value)); label.setStyleSheet("font-size: 28px; font-weight: bold; color:#17365D;")
            bl.addWidget(label); cards.addWidget(box)
        layout.addLayout(cards)

        charts=QHBoxLayout()
        obs_types={r["obs_type"] or "Other": r["c"] for r in db.fetchall("SELECT obs_type,COUNT(*) c FROM observations GROUP BY obs_type")}
        cats={r["category"] or "Other": r["c"] for r in db.fetchall("SELECT category,COUNT(*) c FROM observations GROUP BY category ORDER BY c DESC LIMIT 12")}
        charts.addWidget(PieChartWidget(obs_types),1)
        charts.addWidget(BarChartWidget(cats),1)
        layout.addLayout(charts,1)

        layout.addWidget(QLabel(f"Company: {db.setting('company_name','')}    Project: {db.setting('project_name','')}"))

        table=QTableWidget(); table.setColumnCount(4); table.setHorizontalHeaderLabels(["Observation","Type","Priority","Status"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        rows=db.fetchall("SELECT number,obs_type,priority,status FROM observations ORDER BY id DESC LIMIT 10")
        table.setRowCount(len(rows))
        for r,row in enumerate(rows):
            for c,key in enumerate(["number","obs_type","priority","status"]): table.setItem(r,c,QTableWidgetItem(safe(row[key])))
        layout.addWidget(QLabel("Recent Observations")); layout.addWidget(table)

    def observations(self):

        w, layout = self.page(
            "HSE Inspection Register"
        )

        toolbar = QHBoxLayout()

        search = QLineEdit()
        search.setPlaceholderText(
            "Search number, location, responsible person..."
        )

        toolbar.addWidget(search)

        add_button = QPushButton(
            "+ New Observation"
        )

        toolbar.addWidget(add_button)

        export_button = QPushButton(
            "Export CSV"
        )

        toolbar.addWidget(export_button)

        layout.addLayout(toolbar)

        table = QTableWidget()

        headers = [
            "ID",
            "Number",
            "Date",
            "Type",
            "Category",
            "Location",
            "Responsible",
            "Priority",
            "Target",
            "Status",
            "Overdue"
        ]

        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )
        table.setSelectionBehavior(
            QAbstractItemView.SelectRows
        )

        layout.addWidget(table)

        def load():

            term = search.text().strip()

            query = """
            SELECT * FROM observations
            """

            params = ()

            if term:
                query += """
                WHERE number LIKE ?
                OR location LIKE ?
                OR responsible LIKE ?
                OR observation LIKE ?
                """
                like = f"%{term}%"
                params = (like, like, like, like)

            query += " ORDER BY id DESC"

            rows = db.fetchall(query, params)

            table.setRowCount(len(rows))

            for r, row in enumerate(rows):

                values = [
                    row["id"],
                    row["number"],
                    row["obs_date"],
                    row["obs_type"],
                    row["category"],
                    row["location"],
                    row["responsible"],
                    row["priority"],
                    row["target_date"],
                    row["status"],
                    "OVERDUE"
                    if overdue(
                        row["target_date"],
                        row["status"]
                    )
                    else ""
                ]

                for c, value in enumerate(values):
                    item = QTableWidgetItem(
                        safe(value)
                    )

                    if c == 10 and value:
                        item.setForeground(
                            Qt.GlobalColor.red
                        )

                    table.setItem(r, c, item)

        search.textChanged.connect(load)

        add_button.clicked.connect(
            lambda: self.observation_form(load)
        )

        export_button.clicked.connect(
            lambda: self.export_table(
                table,
                "HSE_Observations"
            )
        )

        load()

    # --------------------------------------------------------

    def observation_form(self, refresh):

        dialog=QDialog(self); dialog.setWindowTitle("New HSE Observation"); dialog.resize(900,780); dialog.setMinimumSize(700,550)
        outer=QVBoxLayout(dialog)
        scroll=QScrollArea(); scroll.setWidgetResizable(True); content=QWidget(); form=QFormLayout(content); scroll.setWidget(content); outer.addWidget(scroll)
        edits={}
        def add_edit(name):
            e=QLineEdit(); edits[name]=e; form.addRow(name.replace("_"," ").title()+":",e); return e
        for name in ["project","company","location","area","responsible","designation","employee_id","observer","observer_designation","observer_id","observer_company"]: add_edit(name)
        obs_type=QComboBox(); obs_type.addItems(OBS_TYPES); form.addRow("Observation Type:",obs_type)
        category=QComboBox(); category.setEditable(True); category.addItems(CATEGORIES); form.addRow("Category:",category)
        subcat=add_edit("subcategory")
        observation=QTextEdit(); observation.setMinimumHeight(90); form.addRow("Observation Description:",observation)
        immediate=QTextEdit(); immediate.setMinimumHeight(70); form.addRow("Immediate Action:",immediate)
        corrective=QTextEdit(); corrective.setMinimumHeight(70); form.addRow("Corrective Action:",corrective)
        preventive=QTextEdit(); preventive.setMinimumHeight(70); form.addRow("Preventive Action:",preventive)
        priority=QComboBox(); priority.addItems(PRIORITIES); form.addRow("Priority:",priority)
        target=QDateEdit(); target.setCalendarPopup(True); target.setDate(datetime.now().date()); form.addRow("Target Completion Date:",target)
        status=QComboBox(); status.addItems(STATUSES); form.addRow("Status:",status)
        closeout=QTextEdit(); closeout.setMinimumHeight(60); form.addRow("Closeout Comments:",closeout)
        evidence={"Observation Evidence":"","Closeout Evidence":""}
        def attach(kind,button):
            path,_=QFileDialog.getOpenFileName(dialog,"Select Evidence Photo","","Images (*.png *.jpg *.jpeg *.bmp *.webp)")
            if path: evidence[kind]=path; button.setText(Path(path).name)
        for kind in evidence:
            b=QPushButton(f"Attach {kind}"); b.clicked.connect(lambda checked=False,k=kind,btn=b:attach(k,btn)); form.addRow(kind+":",b)
        buttons=QDialogButtonBox(QDialogButtonBox.Save|QDialogButtonBox.Cancel); outer.addWidget(buttons); buttons.rejected.connect(dialog.reject)
        def save():
            if not edits["location"].text().strip() or not observation.toPlainText().strip():
                QMessageBox.warning(dialog,"Required","Location and Observation Description are required."); return
            try:
                number=next_number("HSE-OBS","observations")
                db.execute("""INSERT INTO observations(number,obs_date,obs_time,project,company,location,area,responsible,designation,employee_id,observer,observer_designation,observer_id,observer_company,obs_type,category,subcategory,observation,immediate_action,corrective_action,preventive_action,priority,target_date,status,closeout_comments,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(
                    number,today(),now_time(),edits["project"].text(),edits["company"].text(),edits["location"].text(),edits["area"].text(),edits["responsible"].text(),edits["designation"].text(),edits["employee_id"].text(),edits["observer"].text(),edits["observer_designation"].text(),edits["observer_id"].text(),edits["observer_company"].text(),obs_type.currentText(),category.currentText(),subcat.text(),observation.toPlainText(),immediate.toPlainText(),corrective.toPlainText(),preventive.toPlainText(),priority.currentText(),target.date().toString("yyyy-MM-dd"),status.currentText(),closeout.toPlainText(),datetime.now().isoformat()))
                row=db.fetchone("SELECT id FROM observations WHERE number=?",(number,))
                if row:
                    for kind,path in evidence.items():
                        if path:
                            source=Path(path); dest=ATTACH_DIR/f"{number}_{kind.replace(' ','_')}_{source.name}"; shutil.copy2(source,dest)
                            db.execute("INSERT INTO observation_attachments(observation_id,file_path,attachment_type) VALUES(?,?,?)",(row["id"],str(dest),kind))
                dialog.accept(); refresh(); self.dashboard()
            except Exception as e:
                logging.exception("Observation save failed"); QMessageBox.critical(dialog,"Save Error",str(e))
        buttons.accepted.connect(save); dialog.exec()

    def incidents(self):

        w, layout = self.page(
            "Accident / Incident Investigation"
        )

        toolbar = QHBoxLayout()

        add_button = QPushButton(
            "+ New Incident"
        )

        toolbar.addWidget(add_button)

        layout.addLayout(toolbar)

        table = QTableWidget()

        headers = [
            "Number",
            "Date",
            "Type",
            "Location",
            "Project",
            "Method",
            "Status"
        ]

        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )

        layout.addWidget(table)

        def load():

            rows = db.fetchall("""
                SELECT * FROM incidents
                ORDER BY id DESC
            """)

            table.setRowCount(len(rows))

            for r, row in enumerate(rows):

                values = [
                    row["number"],
                    row["incident_date"],
                    row["incident_type"],
                    row["location"],
                    row["project"],
                    row["investigation_method"],
                    row["status"]
                ]

                for c, value in enumerate(values):
                    table.setItem(
                        r, c,
                        QTableWidgetItem(
                            safe(value)
                        )
                    )

        add_button.clicked.connect(
            lambda: self.incident_form(load)
        )

        load()

    # --------------------------------------------------------

    def incident_form(self, refresh):

        dialog=QDialog(self); dialog.setWindowTitle("New Incident Investigation"); dialog.resize(1000,800); dialog.setMinimumSize(760,600)
        outer=QVBoxLayout(dialog); scroll=QScrollArea(); scroll.setWidgetResizable(True); content=QWidget(); main=QVBoxLayout(content); scroll.setWidget(content); outer.addWidget(scroll)
        common=QGroupBox("Incident Information"); cf=QFormLayout(common); edits={}
        def efield(name):
            e=QLineEdit(); edits[name]=e; cf.addRow(name.replace("_"," ").title()+":",e)
        for name in ["incident_time","location","project","company","department","activity","person_involved","employee_id","designation","supervisor","witnesses","equipment"]: efield(name)
        incident_type=QComboBox(); incident_type.addItems(INCIDENT_TYPES); cf.addRow("Incident Type:",incident_type)
        description=QTextEdit(); description.setMinimumHeight(100); cf.addRow("Detailed Incident Description:",description)
        immediate=QTextEdit(); cf.addRow("Immediate Actions Taken:",immediate)
        consequences=QTextEdit(); cf.addRow("Actual Consequences:",consequences)
        potential=QTextEdit(); cf.addRow("Potential Consequences:",potential)
        status=QComboBox(); status.addItems(STATUSES); cf.addRow("Status:",status)
        main.addWidget(common)
        method=QComboBox(); method.addItems(INVESTIGATION_METHODS); main.addWidget(QLabel("Investigation Method")); main.addWidget(method)
        pages=QStackedWidget(); main.addWidget(pages)
        details={}
        def text_page(title, fields):
            box=QWidget(); fl=QFormLayout(box); widgets={}
            for label,key in fields:
                w=QTextEdit(); w.setMinimumHeight(65); widgets[key]=w; fl.addRow(label,w)
            pages.addWidget(box); details[title]=widgets; return widgets
        root=text_page("Root Cause Analysis",[("Problem Statement","problem"),("Immediate / Direct Cause","direct"),("Contributing Causes","contributing"),("Underlying / System Causes","underlying"),("Failed or Missing Controls","controls"),("Human / Organizational Factors","factors"),("Root Cause","root"),("Corrective Action","corrective"),("Preventive Action","preventive"),("Lessons Learned","lessons")])
        whybox=QWidget(); wl=QVBoxLayout(whybox); why_rows=[]
        problem=QTextEdit(); problem.setPlaceholderText("Incident / problem statement"); wl.addWidget(QLabel("Problem / Incident Statement")); wl.addWidget(problem)
        why_container=QVBoxLayout(); wl.addLayout(why_container)
        def add_why():
            n=len(why_rows)+1; row=QHBoxLayout(); q=QLineEdit(); a=QTextEdit(); a.setFixedHeight(50); row.addWidget(QLabel(f"Why {n}")); row.addWidget(q); row.addWidget(QLabel("Answer")); row.addWidget(a); why_container.addLayout(row); why_rows.append((q,a))
        for _ in range(5): add_why()
        addwhy=QPushButton("+ Add Why"); addwhy.clicked.connect(add_why); wl.addWidget(addwhy)
        for label,key in [("Root Cause","root"),("Corrective Action","corrective"),("Preventive Action","preventive"),("Lessons Learned","lessons")]:
            w=QTextEdit(); w.setMinimumHeight(60); why_container.addWidget(QLabel(label)); why_container.addWidget(w); why_rows.append((None,w)) if False else None
            details.setdefault("5 Why Fields",{})[key]=w
        pages.addWidget(whybox)
        fish=QWidget(); fl=QFormLayout(fish); fish_widgets={}
        for cat in ["People","Process","Equipment / Machinery","Materials","Environment","Management / Policy","Training","Supervision","Communication"]:
            w=QTextEdit(); w.setMinimumHeight(55); fish_widgets[cat]=w; fl.addRow(cat,w)
        for label,key in [("Root Cause","root"),("Contributing Causes","contributing"),("Corrective Action","corrective"),("Preventive Action","preventive")]:
            w=QTextEdit(); w.setMinimumHeight(60); fish_widgets[key]=w; fl.addRow(label,w)
        pages.addWidget(fish); details["Fishbone / Ishikawa"]=fish_widgets
        icam=QWidget(); il=QVBoxLayout(icam)
        timeline=QTableWidget(0,5); timeline.setHorizontalHeaderLabels(["Date","Time","Event","Evidence / Source","Person / Remarks"]); timeline.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch); il.addWidget(QLabel("Timeline")); il.addWidget(timeline)
        tb=QPushButton("+ Add Timeline Event"); il.addWidget(tb)
        def add_event():
            r=timeline.rowCount(); timeline.insertRow(r)
            for c in range(5): timeline.setItem(r,c,QTableWidgetItem(""))
        tb.clicked.connect(add_event); add_event()
        icam_fields={}
        for label,key in [("Incident Sequence","sequence"),("Absent / Failed Defences","defences"),("Individual / Team Actions","actions"),("Task / Environmental Conditions","conditions"),("Organizational / Latent Conditions","organizational"),("Findings / Root Causes","findings"),("Corrective Actions","corrective"),("Preventive Actions","preventive"),("Lessons Learned","lessons")]:
            w=QTextEdit(); w.setMinimumHeight(65); icam_fields[key]=w; il.addWidget(QLabel(label)); il.addWidget(w)
        pages.addWidget(icam); details["ICAM"]=icam_fields
        # generic pages for other existing methods
        generic=text_page("Other",[("Analysis / Method Details","details"),("Root Cause / Findings","root"),("Corrective Action","corrective"),("Preventive Action","preventive")]); details["Other"]=generic
        def select_page(i):
            name=method.itemText(i)
            mapping={"Root Cause Analysis":0,"5 Why Analysis":1,"Fishbone / Ishikawa":2,"ICAM":3}
            pages.setCurrentIndex(mapping.get(name,4))
        method.currentIndexChanged.connect(select_page); select_page(method.currentIndex())
        buttons=QDialogButtonBox(QDialogButtonBox.Save|QDialogButtonBox.Cancel); outer.addWidget(buttons); buttons.rejected.connect(dialog.reject)
        def save():
            try:
                number=next_number("HSE-INC","incidents"); method_name=method.currentText(); payload={}
                if method_name=="Root Cause Analysis": payload={k:w.toPlainText() for k,w in root.items()}
                elif method_name=="5 Why Analysis":
                    payload={"problem":problem.toPlainText(),"whys":[{"why":q.text(),"answer":a.toPlainText()} for q,a in why_rows],**{k:w.toPlainText() for k,w in details.get("5 Why Fields",{}).items()}}
                elif method_name=="Fishbone / Ishikawa": payload={k:w.toPlainText() for k,w in fish_widgets.items()}
                elif method_name=="ICAM":
                    events=[]
                    for r in range(timeline.rowCount()): events.append([timeline.item(r,c).text() if timeline.item(r,c) else "" for c in range(5)])
                    payload={"timeline":events,**{k:w.toPlainText() for k,w in icam_fields.items()}}
                else: payload={k:w.toPlainText() for k,w in generic.items()}
                db.execute("""INSERT INTO incidents(number,incident_date,incident_time,location,project,company,department,activity,incident_type,person_involved,employee_id,designation,supervisor,witnesses,description,immediate_action,consequences,potential_consequences,equipment,investigation_method,investigation_details,status,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(number,today(),edits["incident_time"].text(),edits["location"].text(),edits["project"].text(),edits["company"].text(),edits["department"].text(),edits["activity"].text(),incident_type.currentText(),edits["person_involved"].text(),edits["employee_id"].text(),edits["designation"].text(),edits["supervisor"].text(),edits["witnesses"].text(),description.toPlainText(),immediate.toPlainText(),consequences.toPlainText(),potential.toPlainText(),edits["equipment"].text(),method_name,json.dumps(payload,ensure_ascii=False),status.currentText(),datetime.now().isoformat()))
                dialog.accept(); refresh(); self.dashboard()
            except Exception as e:
                logging.exception("Incident save failed"); QMessageBox.critical(dialog,"Save Error",str(e))
        buttons.accepted.connect(save); dialog.exec()

    def audits(self):

        w, layout = self.page(
            "Audit Register"
        )

        toolbar = QHBoxLayout()

        add_button = QPushButton(
            "+ New Audit"
        )

        toolbar.addWidget(add_button)

        layout.addLayout(toolbar)

        table = QTableWidget()

        headers = [
            "Number",
            "Date",
            "Standard",
            "Audit Type",
            "Project",
            "Auditor",
            "Status"
        ]

        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )

        layout.addWidget(table)

        def load():

            rows = db.fetchall("""
                SELECT * FROM audits
                ORDER BY id DESC
            """)

            table.setRowCount(len(rows))

            for r, row in enumerate(rows):

                values = [
                    row["number"],
                    row["audit_date"],
                    row["standard"],
                    row["audit_type"],
                    row["project"],
                    row["lead_auditor"],
                    row["status"]
                ]

                for c, value in enumerate(values):

                    table.setItem(
                        r, c,
                        QTableWidgetItem(
                            safe(value)
                        )
                    )

        add_button.clicked.connect(
            lambda: self.audit_form(load)
        )

        load()

    # --------------------------------------------------------

    def audit_form(self, refresh):

        dialog=QDialog(self); dialog.setWindowTitle("New Audit"); dialog.resize(900,800); dialog.setMinimumSize(720,600)
        outer=QVBoxLayout(dialog); scroll=QScrollArea(); scroll.setWidgetResizable(True); content=QWidget(); form=QFormLayout(content); scroll.setWidget(content); outer.addWidget(scroll)
        edits={}
        def edit(name,label=None):
            w=QLineEdit(); edits[name]=w; form.addRow(label or name.replace("_"," ").title()+":",w); return w
        standard=QComboBox(); standard.addItems(["ISO 45001:2018","ISO 14001:2015"]); form.addRow("Audit Standard:",standard)
        audit_type=QComboBox(); audit_type.addItems(AUDIT_TYPES); form.addRow("Audit Type:",audit_type)
        for name in ["project","location","department","auditor","lead_auditor","auditee","start_time","end_time"]: edit(name)
        scope=QTextEdit(); form.addRow("Scope:",scope); objective=QTextEdit(); form.addRow("Objective:",objective); criteria=QTextEdit(); form.addRow("Audit Criteria:",criteria)
        clause=QComboBox(); sub_clause=QComboBox(); clause_desc=QLabel(); sub_desc=QLabel(); form.addRow("Main Clause:",clause); form.addRow("Clause Description:",clause_desc); form.addRow("Sub-Clause:",sub_clause); form.addRow("Sub-Clause Description:",sub_desc)
        def load_clauses():
            data=ISO_45001_CLAUSES if standard.currentText().startswith("ISO 45001") else ISO_14001_CLAUSES
            clause.blockSignals(True); clause.clear(); clause.addItems(list(data.keys())); clause.blockSignals(False); update_sub()
        def update_sub():
            data=ISO_45001_CLAUSES if standard.currentText().startswith("ISO 45001") else ISO_14001_CLAUSES; key=clause.currentText(); items=data.get(key,[]); sub_clause.clear(); sub_clause.addItems([f"{n} - {d}" for n,d in items]); clause_desc.setText(key.split(" - ",1)[1] if " - " in key else ""); update_sub_desc()
        def update_sub_desc():
            text=sub_clause.currentText(); sub_desc.setText(text.split(" - ",1)[1] if " - " in text else "")
        standard.currentIndexChanged.connect(load_clauses); clause.currentIndexChanged.connect(update_sub); sub_clause.currentIndexChanged.connect(update_sub_desc); load_clauses()
        requirement=QTextEdit(); form.addRow("Requirement / Reference:",requirement)
        finding_type=QComboBox(); finding_type.addItems(AUDIT_FINDING_TYPES); form.addRow("Finding Type:",finding_type)
        observation=QTextEdit(); form.addRow("Observation:",observation); evidence=QTextEdit(); form.addRow("Objective Evidence:",evidence); action=QTextEdit(); form.addRow("Corrective Action:",action)
        responsible=QLineEdit(); form.addRow("Responsible Person:",responsible); target=QLineEdit(); form.addRow("Target Date:",target)
        status=QComboBox(); status.addItems(STATUSES); form.addRow("Status:",status); closeout=QTextEdit(); form.addRow("Closeout Evidence / Verification:",closeout)
        buttons=QDialogButtonBox(QDialogButtonBox.Save|QDialogButtonBox.Cancel); outer.addWidget(buttons); buttons.rejected.connect(dialog.reject)
        def save():
            try:
                number=next_number("HSE-AUD","audits")
                db.execute("""INSERT INTO audits(number,audit_date,audit_type,standard,project,location,department,auditor,lead_auditor,auditee,scope,objective,criteria,start_time,end_time,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(number,today(),audit_type.currentText(),standard.currentText(),edits["project"].text(),edits["location"].text(),edits["department"].text(),edits["auditor"].text(),edits["lead_auditor"].text(),edits["auditee"].text(),scope.toPlainText(),objective.toPlainText(),criteria.toPlainText(),edits["start_time"].text(),edits["end_time"].text(),datetime.now().isoformat()))
                audit=db.fetchone("SELECT id FROM audits WHERE number=?",(number,)); sc=sub_clause.currentText(); scnum=sc.split(" - ",1)[0] if " - " in sc else sc; main_clause=clause.currentText().split(" - ",1)[0]
                db.execute("""INSERT INTO audit_findings(audit_id,clause,sub_clause,requirement,finding_type,observation,evidence,corrective_action,responsible,target_date,status,verification,closeout_evidence) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",(audit["id"],main_clause,scnum,requirement.toPlainText(),finding_type.currentText(),observation.toPlainText(),evidence.toPlainText(),action.toPlainText(),responsible.text(),target.text(),status.currentText(),closeout.toPlainText(),closeout.toPlainText()))
                dialog.accept(); refresh()
            except Exception as e:
                logging.exception("Audit save failed"); QMessageBox.critical(dialog,"Save Error",str(e))
        buttons.accepted.connect(save); dialog.exec()

    def capa(self):

        w, layout = self.page(
            "CAPA Register"
        )

        toolbar = QHBoxLayout()

        add_button = QPushButton(
            "+ New CAPA"
        )

        toolbar.addWidget(add_button)

        layout.addLayout(toolbar)

        table = QTableWidget()

        headers = [
            "Number",
            "Source",
            "Reference",
            "Priority",
            "Responsible",
            "Target Date",
            "Status",
            "Overdue"
        ]

        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )

        layout.addWidget(table)

        def load():

            rows = db.fetchall("""
                SELECT * FROM capa
                ORDER BY id DESC
            """)

            table.setRowCount(len(rows))

            for r, row in enumerate(rows):

                values = [
                    row["number"],
                    row["source"],
                    row["reference_number"],
                    row["priority"],
                    row["responsible"],
                    row["target_date"],
                    row["status"],
                    "OVERDUE"
                    if overdue(
                        row["target_date"],
                        row["status"]
                    )
                    else ""
                ]

                for c, value in enumerate(values):
                    table.setItem(
                        r, c,
                        QTableWidgetItem(
                            safe(value)
                        )
                    )

        add_button.clicked.connect(
            lambda: self.capa_form(load)
        )

        load()

    # --------------------------------------------------------

    def capa_form(self, refresh):

        dialog=QDialog(self); dialog.setWindowTitle("New CAPA"); dialog.resize(850,750); dialog.setMinimumSize(700,550)
        outer=QVBoxLayout(dialog); scroll=QScrollArea(); scroll.setWidgetResizable(True); content=QWidget(); form=QFormLayout(content); scroll.setWidget(content); outer.addWidget(scroll)
        source=QComboBox(); source.addItems(CAPA_SOURCES); form.addRow("Source:",source); reference=QLineEdit(); form.addRow("Reference Number:",reference)
        finding=QTextEdit(); form.addRow("Finding:",finding); root=QTextEdit(); form.addRow("Root Cause:",root); corrective=QTextEdit(); form.addRow("Corrective Action:",corrective); preventive=QTextEdit(); form.addRow("Preventive Action:",preventive)
        responsible=QLineEdit(); form.addRow("Responsible Person:",responsible); priority=QComboBox(); priority.addItems(PRIORITIES); form.addRow("Priority:",priority); target=QLineEdit(); form.addRow("Target Date:",target); status=QComboBox(); status.addItems(STATUSES); form.addRow("Status:",status); verification=QTextEdit(); form.addRow("Verification:",verification); evidence=QTextEdit(); form.addRow("Closeout Evidence:",evidence)
        buttons=QDialogButtonBox(QDialogButtonBox.Save|QDialogButtonBox.Cancel); outer.addWidget(buttons); buttons.rejected.connect(dialog.reject)
        def save():
            try:
                number=next_number("HSE-CAPA","capa")
                db.execute("""INSERT INTO capa(number,source,reference_number,finding,root_cause,corrective_action,preventive_action,responsible,priority,target_date,status,verification,closeout_evidence,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(number,source.currentText(),reference.text(),finding.toPlainText(),root.toPlainText(),corrective.toPlainText(),preventive.toPlainText(),responsible.text(),priority.currentText(),target.text(),status.currentText(),verification.toPlainText(),evidence.toPlainText(),datetime.now().isoformat()))
                dialog.accept(); refresh()
            except Exception as e:
                logging.exception("CAPA save failed"); QMessageBox.critical(dialog,"Save Error",str(e))
        buttons.accepted.connect(save); dialog.exec()

    def reports(self):

        w, layout = self.page(
            "Reports & Export"
        )

        info = QLabel(
            "Generate professional PDF, Excel and CSV reports."
        )

        layout.addWidget(info)

        for title, table_name in [
            ("HSE Observation Report", "observations"),
            ("Incident Report", "incidents"),
            ("Audit Report", "audits"),
            ("CAPA Report", "capa")
        ]:

            box = QGroupBox(title)
            row = QHBoxLayout(box)

            csv_button = QPushButton("CSV")
            excel_button = QPushButton("Excel")
            pdf_button = QPushButton("PDF")

            row.addWidget(csv_button)
            row.addWidget(excel_button)
            row.addWidget(pdf_button)

            csv_button.clicked.connect(
                lambda checked=False,
                t=table_name:
                self.export_csv(t)
            )

            excel_button.clicked.connect(
                lambda checked=False,
                t=table_name:
                self.export_excel(t)
            )

            pdf_button.clicked.connect(
                lambda checked=False,
                t=table_name:
                self.export_pdf(t)
            )

            layout.addWidget(box)

    # --------------------------------------------------------

    def table_data(self, table_name):

        rows = db.fetchall(
            f"SELECT * FROM {table_name} ORDER BY id DESC"
        )

        if not rows:
            return [], []

        columns = rows[0].keys()

        data = []

        for row in rows:
            data.append([
                safe(row[c])
                for c in columns
            ])

        return list(columns), data

    # --------------------------------------------------------

    def export_csv(self, table_name):

        columns, data = self.table_data(
            table_name
        )

        if not columns:
            QMessageBox.information(
                self,
                "No Data",
                "There is no data to export."
            )
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save CSV",
            f"{table_name}_report.csv",
            "CSV (*.csv)"
        )

        if not path:
            return

        with open(
            path,
            "w",
            newline="",
            encoding="utf-8-sig"
        ) as f:

            writer = csv.writer(f)
            writer.writerow(columns)
            writer.writerows(data)

        QMessageBox.information(
            self,
            "Export Complete",
            "CSV report created successfully."
        )

    # --------------------------------------------------------

    def export_excel(self, table_name):

        columns, data = self.table_data(
            table_name
        )

        if not columns:
            QMessageBox.information(
                self,
                "No Data",
                "There is no data to export."
            )
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Excel",
            f"{table_name}_report.xlsx",
            "Excel (*.xlsx)"
        )

        if not path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = table_name[:31]

        ws.append(columns)

        for row in data:
            ws.append(row)

        for cell in ws[1]:
            cell.font = cell.font.copy(
                bold=True
            )

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        wb.save(path)

        QMessageBox.information(
            self,
            "Export Complete",
            "Excel report created successfully."
        )

    # --------------------------------------------------------

    def export_pdf(self, table_name):

        columns, data = self.table_data(
            table_name
        )

        if not columns:
            QMessageBox.information(
                self,
                "No Data",
                "There is no data to export."
            )
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save PDF",
            f"{table_name}_report.pdf",
            "PDF (*.pdf)"
        )

        if not path:
            return

        styles = getSampleStyleSheet()

        doc = SimpleDocTemplate(
            path,
            pagesize=landscape(A4),
            rightMargin=20,
            leftMargin=20,
            topMargin=25,
            bottomMargin=25
        )

        story = []

        company = db.setting(
            "company_name",
            ""
        )

        project = db.setting(
            "project_name",
            ""
        )

        story.append(
            Paragraph(
                company or APP_NAME,
                styles["Title"]
            )
        )

        story.append(
            Paragraph(
                project,
                styles["Heading2"]
            )
        )

        story.append(
            Paragraph(
                f"{table_name.title()} Report",
                styles["Heading2"]
            )
        )

        story.append(
            Paragraph(
                f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                styles["Normal"]
            )
        )

        story.append(Spacer(1, 12))

        # Keep PDF readable by limiting columns.
        max_columns = min(len(columns), 10)

        pdf_columns = columns[:max_columns]

        pdf_data = [pdf_columns]

        for row in data[:500]:

            pdf_data.append(
                row[:max_columns]
            )

        table = Table(
            pdf_data,
            repeatRows=1
        )

        table.setStyle(
            TableStyle([
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#17365D")
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.4,
                    colors.grey
                ),
                (
                    "FONTSIZE",
                    (0, 0),
                    (-1, -1),
                    6
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP"
                )
            ])
        )

        story.append(table)

        footer = db.setting(
            "report_footer",
            ""
        )

        if footer:

            story.append(
                Spacer(1, 10)
            )

            story.append(
                Paragraph(
                    footer,
                    styles["Normal"]
                )
            )

        doc.build(story)

        QMessageBox.information(
            self,
            "Export Complete",
            "PDF report created successfully."
        )

    # --------------------------------------------------------

    def export_table(
        self,
        table,
        filename
    ):

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export CSV",
            f"{filename}.csv",
            "CSV (*.csv)"
        )

        if not path:
            return

        with open(
            path,
            "w",
            newline="",
            encoding="utf-8-sig"
        ) as f:

            writer = csv.writer(f)

            headers = []

            for c in range(
                table.columnCount()
            ):
                headers.append(
                    table.horizontalHeaderItem(
                        c
                    ).text()
                )

            writer.writerow(headers)

            for r in range(
                table.rowCount()
            ):

                row = []

                for c in range(
                    table.columnCount()
                ):

                    item = table.item(r, c)

                    row.append(
                        item.text()
                        if item
                        else ""
                    )

                writer.writerow(row)

        QMessageBox.information(
            self,
            "Export Complete",
            "CSV exported successfully."
        )

    # ========================================================
    # MASTER DATA
    # ========================================================

    def master_data(self):

        w, layout = self.page(
            "Master Data"
        )

        layout.addWidget(
            QLabel(
                "Manage employees, companies, projects and categories."
            )
        )

        for title, table_name in [
            ("Employees", "employees"),
            ("Companies", "companies"),
            ("Projects / Locations", "projects"),
            ("Observation Categories", "categories")
        ]:

            button = QPushButton(title)

            button.clicked.connect(
                lambda checked=False,
                t=table_name,
                n=title:
                self.master_editor(
                    t,
                    n
                )
            )

            layout.addWidget(button)

    # --------------------------------------------------------

    def master_editor(
        self,
        table_name,
        title
    ):

        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(650, 550)

        layout = QVBoxLayout(dialog)

        table = QTableWidget()

        layout.addWidget(table)

        form = QFormLayout()

        if table_name == "employees":

            fields = [
                ("Employee ID", "employee_id"),
                ("Name", "name"),
                ("Designation", "designation"),
                ("Company", "company"),
                ("Department", "department"),
                ("Contact", "contact")
            ]

        elif table_name == "companies":

            fields = [
                ("Company Name", "name"),
                ("Company Type", "company_type"),
                ("Contact Person", "contact_person"),
                ("Contact Number", "contact_number")
            ]

        elif table_name == "projects":

            fields = [
                ("Project", "project"),
                ("Area", "area"),
                ("Location", "location"),
                ("Description", "description")
            ]

        else:

            fields = [
                ("Category", "category"),
                ("Subcategory", "subcategory")
            ]

        inputs = {}

        for label, key in fields:

            e = QLineEdit()

            inputs[key] = e

            form.addRow(
                label,
                e
            )

        layout.addLayout(form)

        add = QPushButton("Add")
        layout.addWidget(add)

        close = QPushButton("Close")
        layout.addWidget(close)

        close.clicked.connect(
            dialog.accept
        )

        def load():

            rows = db.fetchall(
                f"SELECT * FROM {table_name}"
            )

            if not rows:
                table.setRowCount(0)
                return

            columns = list(rows[0].keys())

            table.setColumnCount(
                len(columns)
            )

            table.setHorizontalHeaderLabels(
                columns
            )

            table.setRowCount(
                len(rows)
            )

            for r, row in enumerate(rows):

                for c, col in enumerate(columns):

                    table.setItem(
                        r,
                        c,
                        QTableWidgetItem(
                            safe(row[col])
                        )
                    )

            table.horizontalHeader().setSectionResizeMode(
                QHeaderView.Stretch
            )

        def save():

            values = [
                inputs[key].text()
                for _, key in fields
            ]

            if not values[0].strip():
                return

            placeholders = ",".join(
                ["?"] * len(values)
            )

            columns = ",".join(
                [key for _, key in fields]
            )

            db.execute(
                f"""
                INSERT INTO {table_name}
                ({columns})
                VALUES ({placeholders})
                """,
                values
            )

            for e in inputs.values():
                e.clear()

            load()

        add.clicked.connect(save)

        load()

        dialog.exec()

    # ========================================================
    # SETTINGS
    # ========================================================

    def settings_page(self):

        w, layout = self.page("Settings")
        form=QFormLayout()
        company=QLineEdit(db.setting("company_name","")); project=QLineEdit(db.setting("project_name","")); location=QLineEdit(db.setting("default_location","")); observer=QLineEdit(db.setting("default_observer","")); footer=QLineEdit(db.setting("report_footer",""))
        for label,widget in [("Company Name",company),("Project Name",project),("Default Location",location),("Default Observer",observer),("Report Footer",footer)]: form.addRow(label,widget)
        layout.addLayout(form)
        sap=QGroupBox("SAP"); sf=QFormLayout(sap); sap_enabled=QComboBox(); sap_enabled.addItems(["Disabled","Enabled"]); sap_enabled.setCurrentText(db.setting("sap_enabled","Disabled")); sap_server=QLineEdit(db.setting("sap_server","")); sap_client=QLineEdit(db.setting("sap_client","")); sap_env=QLineEdit(db.setting("sap_environment","")); sf.addRow("SAP Integration",sap_enabled); sf.addRow("SAP Server",sap_server); sf.addRow("SAP Client",sap_client); sf.addRow("SAP Environment",sap_env); layout.addWidget(sap)
        save=QPushButton("Save Settings"); layout.addWidget(save)
        def save_all():
            self.save_settings(company,project,location,observer,footer); db.set_setting("sap_enabled",sap_enabled.currentText()); db.set_setting("sap_server",sap_server.text()); db.set_setting("sap_client",sap_client.text()); db.set_setting("sap_environment",sap_env.text())
        save.clicked.connect(save_all)
        backup=QPushButton("Backup Complete Application Data"); layout.addWidget(backup); backup.clicked.connect(self.backup)
        restore=QPushButton("Restore Application Backup"); layout.addWidget(restore); restore.clicked.connect(self.restore)
        about=QPushButton("About"); layout.addWidget(about); about.clicked.connect(lambda:QMessageBox.information(self,"About",f"{APP_NAME}\nVersion {APP_VERSION}\n\nOffline HSE Management System"))

    def backup(self):

        destination, _ = QFileDialog.getSaveFileName(
            self,
            "Save HSE Backup",
            f"HSE_Backup_{today()}.zip",
            "ZIP (*.zip)"
        )

        if not destination:
            return

        try:

            with zipfile.ZipFile(
                destination,
                "w",
                zipfile.ZIP_DEFLATED
            ) as z:

                if DB_FILE.exists():

                    z.write(
                        DB_FILE,
                        "database/hse.db"
                    )

                for file in ATTACH_DIR.rglob("*"):

                    if file.is_file():

                        z.write(
                            file,
                            f"attachments/{file.name}"
                        )

            QMessageBox.information(
                self,
                "Backup Complete",
                "HSE backup created successfully."
            )

        except Exception as e:

            logging.exception(
                "Backup failed"
            )

            QMessageBox.critical(
                self,
                "Backup Error",
                str(e)
            )

    # --------------------------------------------------------

    def restore(self):

        source, _ = QFileDialog.getOpenFileName(
            self,
            "Select HSE Backup",
            "",
            "ZIP (*.zip)"
        )

        if not source:
            return

        answer = QMessageBox.question(
            self,
            "Restore Backup",
            "Restoring will replace the current database.\n\n"
            "Do you want to continue?"
        )

        if answer != QMessageBox.Yes:
            return

        try:

            # Close current database connection.
            db.conn.close()

            with zipfile.ZipFile(
                source,
                "r"
            ) as z:

                temp = APP_DIR / "restore_temp"

                if temp.exists():
                    shutil.rmtree(temp)

                temp.mkdir()

                z.extractall(temp)

                restored_db = (
                    temp /
                    "database" /
                    "hse.db"
                )

                if not restored_db.exists():
                    raise Exception(
                        "Backup database not found."
                    )

                shutil.copy2(
                    restored_db,
                    DB_FILE
                )

                restored_attach = (
                    temp /
                    "attachments"
                )

                if restored_attach.exists():

                    for file in restored_attach.iterdir():

                        if file.is_file():

                            shutil.copy2(
                                file,
                                ATTACH_DIR /
                                file.name
                            )

                shutil.rmtree(temp)

            QMessageBox.information(
                self,
                "Restore Complete",
                "Backup restored successfully.\n"
                "Please restart the application."
            )

        except Exception as e:

            logging.exception(
                "Restore failed"
            )

            QMessageBox.critical(
                self,
                "Restore Error",
                str(e)
            )

        finally:

            # Reopen database.
            db.conn = sqlite3.connect(
                DB_FILE
            )

            db.conn.row_factory = sqlite3.Row
            db.conn.execute(
                "PRAGMA foreign_keys = ON"
            )

    # ========================================================
    # CLOSE
    # ========================================================

    def closeEvent(self, event):

        try:
            db.conn.commit()
            db.conn.close()
        except Exception:
            pass

        event.accept()


# ============================================================
# APPLICATION START
# ============================================================

def main():

    app = QApplication(sys.argv)

    app.setApplicationName(
        APP_NAME
    )

    app.setOrganizationName(
        "HSE"
    )

    window = MainWindow()

    window.show()

    sys.exit(
        app.exec()
    )


if __name__ == "__main__":
    main()
