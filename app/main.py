import sys
import os
import csv
import shutil
import sqlite3
import zipfile
import logging
import json
from copy import copy
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
    QHeaderView, QAbstractItemView, QCheckBox, QScrollArea
)

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
)
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, PatternFill, Border, Side
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


APP_NAME = "HSE Management System"
APP_VERSION = "1.4.8"

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

        CREATE TABLE IF NOT EXISTS incident_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            incident_id INTEGER NOT NULL,
            file_path TEXT,
            attachment_type TEXT DEFAULT 'Evidence',
            remark TEXT DEFAULT '',
            FOREIGN KEY(incident_id)
                REFERENCES incidents(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS audit_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            audit_id INTEGER NOT NULL,
            file_path TEXT,
            attachment_type TEXT DEFAULT 'Evidence',
            FOREIGN KEY(audit_id)
                REFERENCES audits(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS capa_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            capa_id INTEGER NOT NULL,
            file_path TEXT,
            attachment_type TEXT DEFAULT 'Evidence',
            FOREIGN KEY(capa_id)
                REFERENCES capa(id)
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

        # Backward-compatible migrations for existing installations.
        self.conn.execute("CREATE TABLE IF NOT EXISTS incident_attachments (id INTEGER PRIMARY KEY AUTOINCREMENT, incident_id INTEGER NOT NULL, file_path TEXT, attachment_type TEXT DEFAULT 'Evidence', FOREIGN KEY(incident_id) REFERENCES incidents(id) ON DELETE CASCADE)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS audit_attachments (id INTEGER PRIMARY KEY AUTOINCREMENT, audit_id INTEGER NOT NULL, file_path TEXT, attachment_type TEXT DEFAULT 'Evidence', FOREIGN KEY(audit_id) REFERENCES audits(id) ON DELETE CASCADE)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS capa_attachments (id INTEGER PRIMARY KEY AUTOINCREMENT, capa_id INTEGER NOT NULL, file_path TEXT, attachment_type TEXT DEFAULT 'Evidence', FOREIGN KEY(capa_id) REFERENCES capa(id) ON DELETE CASCADE)")
        # Newer incident-investigation procedure data is stored as JSON so existing
        # installations keep all original columns/features without a destructive migration.
        incident_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(incidents)").fetchall()}
        if "investigation_details" not in incident_cols:
            self.conn.execute("ALTER TABLE incidents ADD COLUMN investigation_details TEXT DEFAULT ''")
        # Incident/RCA extensions are additive only; existing data remains intact.
        incident_extra = {
            "incident_title": "TEXT DEFAULT ''",
            "people_details": "TEXT DEFAULT ''",
            "equipment_details": "TEXT DEFAULT ''",
            "environmental_category": "TEXT DEFAULT ''",
            "environmental_quantity": "TEXT DEFAULT ''",
            "property_name": "TEXT DEFAULT ''",
            "property_damage_cost": "TEXT DEFAULT ''",
            "organization": "TEXT DEFAULT ''",
            "procedure_reference": "TEXT DEFAULT ''",
            "report_reference": "TEXT DEFAULT ''"
        }
        for col, definition in incident_extra.items():
            if col not in incident_cols:
                self.conn.execute(f"ALTER TABLE incidents ADD COLUMN {col} {definition}")
        # Add attachment remarks without replacing any existing attachment data.
        incident_attachment_cols = [r[1] for r in self.conn.execute("PRAGMA table_info(incident_attachments)").fetchall()]
        if "remark" not in incident_attachment_cols:
            self.conn.execute("ALTER TABLE incident_attachments ADD COLUMN remark TEXT DEFAULT ''")
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


def xml_safe(value):
    """Remove characters that are illegal in XML/Office Open XML files."""
    text = safe(value)
    return "".join(ch for ch in text if ch in "\t\n\r" or ord(ch) >= 32)


def attachment_names(table_name, record_id):
    return "; ".join(Path(safe(r["file_path"])).name for r in attachment_rows(table_name, record_id))


def attachment_export_map(table_name):
    return {
        "observations": "observation_attachments",
        "incidents": "incident_attachments",
        "audits": "audit_attachments",
        "capa": "capa_attachments",
    }.get(table_name)


def overdue(target, status):
    if not target or status in ("Closed", "Cancelled"):
        return False
    try:
        return date.fromisoformat(target) < date.today()
    except Exception:
        return False



# ============================================================
# DOCUMENT / ATTACHMENT HELPERS
# ============================================================

ISO_CLAUSES = {
    "ISO 45001": {
        "4": ["4.1 Understanding the organization and its context",
              "4.2 Understanding the needs and expectations of workers and other interested parties",
              "4.3 Determining the scope of the OH&S management system",
              "4.4 OH&S management system"],
        "5": ["5.1 Leadership and commitment", "5.2 OH&S policy",
              "5.3 Organizational roles, responsibilities and authorities",
              "5.4 Consultation and participation of workers"],
        "6": ["6.1 Actions to address risks and opportunities",
              "6.1.2 Hazard identification and assessment of risks and opportunities",
              "6.1.3 Determination of legal requirements and other requirements",
              "6.2 OH&S objectives and planning to achieve them"],
        "7": ["7.1 Resources", "7.2 Competence", "7.3 Awareness",
              "7.4 Communication", "7.5 Documented information"],
        "8": ["8.1 Operational planning and control", "8.1.2 Eliminating hazards and reducing OH&S risks",
              "8.2 Emergency preparedness and response"],
        "9": ["9.1 Monitoring, measurement, analysis and performance evaluation",
              "9.2 Internal audit", "9.3 Management review"],
        "10": ["10.1 General", "10.2 Incident, nonconformity and corrective action",
               "10.3 Continual improvement"],
    },
    "ISO 14001": {
        "4": ["4.1 Understanding the organization and its context",
              "4.2 Understanding the needs and expectations of interested parties",
              "4.3 Determining the scope of the environmental management system",
              "4.4 Environmental management system"],
        "5": ["5.1 Leadership and commitment", "5.2 Environmental policy",
              "5.3 Organizational roles, responsibilities and authorities"],
        "6": ["6.1 Actions to address risks and opportunities",
              "6.1.2 Environmental aspects", "6.1.3 Compliance obligations",
              "6.1.4 Planning action", "6.2 Environmental objectives and planning"],
        "7": ["7.1 Resources", "7.2 Competence", "7.3 Awareness",
              "7.4 Communication", "7.5 Documented information"],
        "8": ["8.1 Operational planning and control", "8.2 Emergency preparedness and response"],
        "9": ["9.1 Monitoring, measurement, analysis and evaluation",
              "9.2 Internal audit", "9.3 Management review"],
        "10": ["10.1 General", "10.2 Nonconformity and corrective action",
               "10.3 Continual improvement"],
    }
}


def company_name():
    return db.setting("company_name", "") or APP_NAME


def document_prefix():
    return db.setting("document_prefix", "HSE") or "HSE"


def document_number(kind):
    prefix = document_prefix()
    return next_number(f"{prefix}-{kind}", {
        "INC": "incidents",
        "OBS": "observations",
        "AUD": "audits",
        "CAPA": "capa"
    }.get(kind, "incidents"))


def attachment_rows(table_name, record_id):
    allowed = {
        "observation_attachments": "observation_id",
        "incident_attachments": "incident_id",
        "audit_attachments": "audit_id",
        "capa_attachments": "capa_id"
    }
    if table_name not in allowed:
        return []
    return db.fetchall(
        f"SELECT * FROM {table_name} WHERE {allowed[table_name]}=? ORDER BY id",
        (record_id,)
    )


def copy_attachments(paths, record_number, table_name, record_id, attachment_type="Evidence"):
    column = {
        "observation_attachments": "observation_id",
        "incident_attachments": "incident_id",
        "audit_attachments": "audit_id",
        "capa_attachments": "capa_id"
    }[table_name]
    saved = []
    for path in paths or []:
        try:
            source = Path(path)
            if not source.exists():
                continue
            destination = ATTACH_DIR / f"{record_number}_{source.name}"
            counter = 1
            while destination.exists():
                destination = ATTACH_DIR / f"{record_number}_{counter}_{source.name}"
                counter += 1
            shutil.copy2(source, destination)
            db.execute(
                f"INSERT INTO {table_name} ({column},file_path,attachment_type) VALUES(?,?,?)",
                (record_id, str(destination), attachment_type)
            )
            saved.append(str(destination))
        except Exception:
            logging.exception("Failed to save attachment %s", path)
    return saved


def add_docx_cell_shading(cell, fill="17365D"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_docx_cell_borders(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge in kwargs:
            edge_data = kwargs.get(edge)
            tag = "w:{}".format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["val", "sz", "space", "color"]:
                if key in edge_data:
                    element.set(qn("w:{}".format(key)), str(edge_data[key]))


def style_docx_table(table, header=True, font_size=7):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False
    for row_index, row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_docx_cell_borders(
                cell,
                top={"val": "single", "sz": 6, "color": "808080"},
                bottom={"val": "single", "sz": 6, "color": "808080"},
                left={"val": "single", "sz": 6, "color": "808080"},
                right={"val": "single", "sz": 6, "color": "808080"},
            )
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
            if header and row_index == 0:
                add_docx_cell_shading(cell)
                for run in cell.paragraphs[0].runs:
                    run.font.bold = True
                    run.font.color.rgb = __import__("docx").shared.RGBColor(255,255,255)


def report_logo_path():
    """Resolve the logo saved by Settings, including relative/legacy paths."""
    configured=safe(db.setting("company_logo",""))
    candidates=[]
    if configured:
        cp=Path(configured)
        candidates.extend([cp, APP_DIR / cp, ATTACH_DIR / cp.name])
    candidates += sorted(ATTACH_DIR.glob("company_logo.*"))
    seen=set()
    for p in candidates:
        try:
            p=Path(p)
            key=str(p.resolve()) if p.exists() else str(p)
            if key in seen: continue
            seen.add(key)
            if p.exists() and p.is_file(): return p
        except Exception:
            pass
    return None

def add_docx_report_header(document,title,kind="REPORT",report_no=""):
    section=document.sections[0]; header=section.header
    usable=(section.page_width-section.left_margin-section.right_margin)/914400.0
    if usable < 9.0:
        widths=[Inches(1.20),Inches(4.05),Inches(2.20)]
    else:
        widths=[Inches(1.55),Inches(5.55),Inches(3.2)]
    table=header.add_table(rows=1,cols=3,width=Inches(sum(w.inches for w in widths))); table.autofit=False
    for cell,w in zip(table.rows[0].cells,widths): cell.width=w; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    prefix=document_prefix(); report_no=report_no or f"{prefix}-{kind}-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}"; doc_no=f"{prefix}-{kind}"; project=db.setting("project_name","")
    left,mid,right=table.rows[0].cells
    lp=left.paragraphs[0]; lp.alignment=WD_ALIGN_PARAGRAPH.CENTER; logo=report_logo_path()
    if logo:
        try: lp.add_run().add_picture(str(logo),width=Inches(1.15))
        except Exception: lp.add_run(xml_safe(company_name())).bold=True
    else: lp.add_run(xml_safe(company_name())).bold=True
    mp=mid.paragraphs[0]; mp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=mp.add_run(xml_safe(title)); r.bold=True; r.font.size=Pt(14)
    r=mp.add_run("\n"+xml_safe(company_name())); r.bold=True; r.font.size=Pt(9)
    rp=right.paragraphs[0]; rp.alignment=WD_ALIGN_PARAGRAPH.LEFT
    for label,val in [("Report No.",report_no),("Document No.",doc_no),("Project",project)]:
        rr=rp.add_run(f"{label}: {xml_safe(val)}\n"); rr.font.size=Pt(8); rr.bold=(label=="Report No.")
    for cell in table.rows[0].cells:
        set_docx_cell_borders(cell,top={"val":"single","sz":8,"color":"808080"},bottom={"val":"single","sz":8,"color":"808080"},left={"val":"single","sz":8,"color":"808080"},right={"val":"single","sz":8,"color":"808080"})
    document.add_paragraph().paragraph_format.space_after=Pt(0)

def add_docx_header(document):
    add_docx_report_header(document,"HSE REPORT")


def add_docx_footer(document):
    footer_text = db.setting("report_footer", "")
    if not footer_text:
        return
    p = document.sections[0].footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(xml_safe(footer_text)).font.size = Pt(8)


def add_docx_title(document, title, number=""):
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(xml_safe(title))
    r.bold = True
    r.font.size = Pt(18)
    if number:
        p2 = document.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(xml_safe(f"Document No.: {number}"))
        r2.bold = True
        r2.font.size = Pt(10)

def save_docx_validated(document, path):
    """Write a Word report atomically and validate the OOXML ZIP before replacing the target."""
    target=Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temp=target.with_name(target.name + ".tmp.docx")
    try:
        document.save(str(temp))
        with zipfile.ZipFile(temp, "r") as z:
            bad=z.testzip()
            if bad:
                raise ValueError(f"Generated Word report is corrupt: {bad}")
        os.replace(temp, target)
    finally:
        temp.unlink(missing_ok=True)



def add_docx_kv_table(document, pairs):
    table = document.add_table(rows=0, cols=2)
    for key, value in pairs:
        cells = table.add_row().cells
        cells[0].text = xml_safe(key)
        cells[1].text = xml_safe(value)
        cells[0].paragraphs[0].runs[0].bold = True
    style_docx_table(table, header=False)
    return table


def add_docx_attachments(document, rows):
    document.add_heading("Evidence / Attachments", level=2)
    if not rows:
        document.add_paragraph("No attachments recorded.")
        return
    for i,row in enumerate(rows,1):
        path=Path(safe(row["file_path"])); remark=safe(row["remark"]) if "remark" in row.keys() else ""
        document.add_paragraph(f"{i}. {path.name} | Type: {safe(row['attachment_type'])} | Remark: {remark}")
        if path.exists() and path.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
            try:
                p=document.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(str(path), width=Inches(5.8))
            except Exception:
                pass
        elif path.exists() and path.suffix.lower()==".pdf":
            document.add_paragraph("PDF attachment: " + path.name)
    

def add_docx_signatures(document):
    document.add_paragraph()
    document.add_heading("Review and Approval", level=2)
    table = document.add_table(rows=2, cols=3)
    labels = ["Prepared / HSE Report Made By", "Reviewed By", "Approved By"]
    for i, label in enumerate(labels):
        table.cell(0, i).text = label
        table.cell(1, i).text = "\n\n____________________________\nSignature: __________________\nDate: ______________________"
    style_docx_table(table, header=True)


def generate_incident_docx(incident_id, path):
    row=db.fetchone("SELECT * FROM incidents WHERE id=?",(incident_id,))
    if not row: raise ValueError("Incident record not found.")
    doc=Document()
    sec=doc.sections[0]
    sec.orientation=WD_ORIENT.PORTRAIT
    sec.page_width=Inches(8.27); sec.page_height=Inches(11.69)
    sec.left_margin=Inches(0.35); sec.right_margin=Inches(0.35); sec.top_margin=Inches(0.35); sec.bottom_margin=Inches(0.35)
    add_docx_report_header(doc,"ACCIDENT / INCIDENT INVESTIGATION REPORT","INC",row["number"])
    title=safe(row["incident_title"]) if "incident_title" in row.keys() else ""
    add_docx_kv_table(doc,[
        ("Incident Title",title),("Incident Date",row["incident_date"]),("Incident Time",row["incident_time"]),
        ("Company",row["company"] or company_name()),("Project",row["project"]),("Location",row["location"]),
        ("Department",row["department"]),("Activity",row["activity"]),("Incident Type",row["incident_type"]),
        ("Investigation Method",row["investigation_method"]),("Status",row["status"]),
        ("Organisation",row["organization"]),("Report / Reference No.",row["report_reference"])
    ])
    try: details=json.loads(safe(row["investigation_details"]) or "{}")
    except Exception: details={}
    def j(col):
        try:return json.loads(safe(row[col]) or "[]")
        except Exception:return []
    def section_table(title, headers, data):
        if not data:return
        doc.add_heading(title,2); t=doc.add_table(rows=1,cols=len(headers))
        for i,h in enumerate(headers): t.rows[0].cells[i].text=h
        for vals in data:
            c=t.add_row().cells
            for i,v in enumerate(vals): c[i].text=xml_safe(v)
        style_docx_table(t,font_size=8)
    people=j("people_details"); equipment=j("equipment_details")
    section_table("People Involved",["Person Involved","Role","Designation","ID"],[[x.get("name",""),x.get("role",""),x.get("designation",""),x.get("id","")] for x in people])
    section_table("Equipment",["Equipment","ID","Specification"],[[x.get("name",""),x.get("id",""),x.get("specification","")] for x in equipment])
    for key,titleh,headers in [("environmental_items","Environmental Information",["Category","Quantity / Unit"]),("property_items","Property / Asset",["Property / Asset Name","Cost of Damage"]),("procedure_items","Procedure / Reference",["Reference","Details"])]:
        items=details.get(key,[]) or []
        vals=[]
        for x in items:
            if key=="environmental_items": vals.append([x.get("category",""),x.get("quantity","")])
            elif key=="property_items": vals.append([x.get("name",""),x.get("cost","")])
            else: vals.append([x.get("reference",""),x.get("details","")])
        section_table(titleh,headers,vals)
    for h,v in [("Incident Description",row["description"]),("Immediate Action",row["immediate_action"])]:
        if v: doc.add_heading(h,2); doc.add_paragraph(xml_safe(v))
    # Witness statements, including optional attachment and remark for each statement.
    witnesses=details.get("witness_statements",[]) or []
    if witnesses:
        doc.add_heading("Witness Statements",2)
        for i,w in enumerate(witnesses,1):
            doc.add_paragraph(f"Witness {i}: {xml_safe(w.get('name',''))} | Remark: {xml_safe(w.get('remark',''))}")
            doc.add_paragraph(xml_safe(w.get("statement","")))
            if w.get("attachment"): doc.add_paragraph("Statement Attachment: "+xml_safe(w.get("attachment")))
    if row["investigation_method"]=="Fishbone / Ishikawa":
        fish=details.get("fishbone",[]) or []
        doc.add_heading("Fishbone / Ishikawa Analysis",2)
        # Keep each Fishbone cause attachment beside the exact cause that it supports.
        if fish:
            incident_attach_map={Path(safe(a["file_path"])).name:Path(safe(a["file_path"])) for a in attachment_rows("incident_attachments",incident_id)}
            ft=doc.add_table(rows=1,cols=5)
            for i,h in enumerate(["Factor","Suggested Cause","Specific Cause / Explanation","Evidence / Finding","Attachment"]): ft.rows[0].cells[i].text=h
            for x in fish:
                cells=ft.add_row().cells
                vals=[x.get("category",""),x.get("cause",""),x.get("detail",""),x.get("evidence","")]
                for i,v in enumerate(vals): cells[i].text=xml_safe(v)
                att_names=x.get("attachments",[]) or []
                if not att_names: cells[4].text="None"
                else:
                    cells[4].text=""
                    for nm in att_names:
                        fp=incident_attach_map.get(Path(nm).name, ATTACH_DIR / Path(nm).name)
                        if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                            try:
                                pp=cells[4].paragraphs[0] if not cells[4].paragraphs[0].runs else cells[4].add_paragraph()
                                pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; pp.add_run().add_picture(str(fp),width=Inches(1.25))
                                cells[4].add_paragraph(xml_safe(fp.name))
                            except Exception: cells[4].add_paragraph(xml_safe(fp.name))
                        else:
                            cells[4].add_paragraph(xml_safe(fp.name))
            style_docx_table(ft,font_size=7)
        team=details.get("investigation_team",[]) or []
        section_table("Investigation Team",["Name","Position","Department","Role in Investigation"],[[x.get("name",""),x.get("position",""),x.get("department",""),x.get("role","")] for x in team])
        timeline=details.get("timeline_events",[]) or []
        section_table("Timeline of Events",["Time","Event / What Happened"],[[x.get("time",""),x.get("event","")] for x in timeline])
        if details.get("problem_statement"):
            doc.add_heading("Problem Statement",2); doc.add_paragraph(xml_safe(details.get("problem_statement")))
        ver=details.get("root_cause_verification",[]) or []
        section_table("Root Cause Verification",["Potential Cause","Evidence / Finding","Verified?","Responsible"],[[x.get("cause",""),x.get("evidence",""),x.get("verified",""),x.get("responsible","")] for x in ver])
        if details.get("effectiveness_verification"):
            doc.add_heading("Effectiveness Verification",2); doc.add_paragraph(xml_safe(details.get("effectiveness_verification")))
        if details.get("lessons_learned"):
            doc.add_heading("Lessons Learned",2); doc.add_paragraph(xml_safe(details.get("lessons_learned")))
        if details.get("conclusion"):
            doc.add_heading("Conclusion",2); doc.add_paragraph(xml_safe(details.get("conclusion")))
        # Cause-and-effect summary showing the Fishbone factors leading to the incident/problem.
        doc.add_heading("Cause-and-Effect Diagram",2)
        factor_groups={}
        for x in fish:
            factor_groups.setdefault(x.get("category","Other"),[]).append(x.get("cause","") or x.get("detail",""))
        diagram_rows=[["Factor","Possible Causes","Effect / Problem"]]
        for cat in ["People","Machine / Equipment","Method","Material","Measurement","Environment","Other"]:
            vals=factor_groups.get(cat,[])
            if vals: diagram_rows.append([cat,"; ".join(v for v in vals if v),safe(row["incident_title"]) or "INCIDENT / PROBLEM"])
        section_table("Fishbone Cause-and-Effect",["Main Factor","Possible Causes","Effect / Problem"],diagram_rows[1:] if len(diagram_rows)>1 else [])
    for h,key,fallback in [("Direct Cause","direct_causes",row["direct_cause"]),("Contributing Factors","contributing_factors_list",row["contributing_factors"]),("Root Cause","root_causes_list",row["root_cause"])]:
        vals=details.get(key,[]); vals=vals if isinstance(vals,list) else ([vals] if vals else [])
        if not vals and fallback: vals=[fallback]
        doc.add_heading(h,2)
        for i,v in enumerate(vals,1): doc.add_paragraph(f"{i}. {xml_safe(v)}")
    if row["investigation_method"]=="5 Why Analysis" or any(row[f"why{i}"] for i in range(1,6)):
        section_table("5 Why Analysis",["Step","Analysis"],[[f"Why {i}",row[f"why{i}"] or "N/A"] for i in range(1,6)])
    actions=details.get("corrective_actions",[]) or []
    section_table("Corrective Action Plan",["Action","Responsible","Target Date","Status","Closeout Evidence"],[[a.get("action",""),a.get("responsible",""),a.get("target_date",""),a.get("status","Open"),"; ".join(a.get("attachments",[])) or "None"] for a in actions])
    add_docx_attachments(doc,attachment_rows("incident_attachments",incident_id))
    add_docx_signatures(doc); add_docx_footer(doc); save_docx_validated(doc,path)


OBSERVATION_EXPORT_FIELDS = [
    ("S.No.", "__serial__"), ("Observation No.", "number"), ("Observation Date", "obs_date"),
    ("Observation Time", "obs_time"), ("Project", "project"), ("Company", "company"),
    ("Location", "location"), ("Area", "area"), ("Responsible Person", "responsible"),
    ("Designation", "designation"), ("Employee ID", "employee_id"), ("Observer", "observer"),
    ("Observer Designation", "observer_designation"), ("Observer ID", "observer_id"),
    ("Observer Company", "observer_company"), ("Observation Type", "obs_type"),
    ("Category", "category"), ("Subcategory", "subcategory"),
    ("Observation Detail", "observation"),
    ("Action Taken", "corrective_action"),
    ("Priority", "priority"), ("Target Date", "target_date"), ("Status", "status"),
    ("Closeout Date", "closeout_date"), ("Closed By", "closed_by"),
    ("Verified By", "verified_by"), ("Verification Date", "verification_date"),
    ("Closeout Comments", "closeout_comments"), ("Verification", "verification"),
    ("Evidence", "__attachments__"),
]


def observation_export_rows(record_ids=None):
    params=[]
    where=""
    if record_ids:
        placeholders=",".join("?" for _ in record_ids)
        where=f" WHERE id IN ({placeholders})"
        params=list(record_ids)
    # Current records only; deleted observations can never contribute to the serial sequence.
    rows=db.fetchall(f"SELECT * FROM observations{where} ORDER BY id DESC", params)
    columns=[x[0] for x in OBSERVATION_EXPORT_FIELDS]
    data=[]
    for serial, row in enumerate(rows, 1):
        values=[]
        for _, key in OBSERVATION_EXPORT_FIELDS:
            if key == "__serial__":
                value=serial
            elif key == "__attachments__":
                value=attachment_names("observation_attachments", row["id"])
            else:
                value=row[key]
            values.append(xml_safe(value))
        data.append(values)
    return columns, data, rows


def observation_word_sequence():
    return [
        "S.No.", "Observation Date", "Location", "Observation Type", "Category",
        "Observation Detail", "Action Taken", "Responsible", "Evidence", "Status"
    ]


def observation_attachment_images(record_id):
    result=[]
    for a in attachment_rows("observation_attachments", record_id):
        p=Path(safe(a["file_path"]))
        if p.exists() and p.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp"}:
            result.append((p, safe(a["attachment_type"])))
    return result

def set_docx_landscape(document):
    section=document.sections[0]
    section.orientation=WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height=section.page_height, section.page_width
    section.left_margin=Inches(0.35)
    section.right_margin=Inches(0.35)
    section.top_margin=Inches(0.45)
    section.bottom_margin=Inches(0.45)


def add_observation_attachments_to_docx(document, observation_rows):
    """Add compact attachment information without creating a second wide table."""
    document.add_heading("Attachments / Evidence", level=2)
    table = document.add_table(rows=1, cols=3)
    hdr = table.rows[0].cells
    hdr[0].text = "S.No."
    hdr[1].text = "Observation No."
    hdr[2].text = "Attachment / Evidence"
    any_attachment = False
    for serial, obs in enumerate(observation_rows, 1):
        attachments = attachment_rows("observation_attachments", obs["id"])
        if not attachments:
            continue
        any_attachment = True
        cells = table.add_row().cells
        cells[0].text = str(serial)
        cells[1].text = safe(obs["number"])
        for idx, a in enumerate(attachments):
            path = Path(safe(a["file_path"]))
            if idx:
                cells[2].add_paragraph()
            p = cells[2].paragraphs[-1]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(path.name)
            run.font.size = Pt(7)
            if path.exists() and path.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp"}:
                try:
                    p.add_run("  ")
                    p.add_run().add_picture(str(path), width=Inches(0.55))
                except Exception:
                    pass
    if not any_attachment:
        cells = table.add_row().cells
        cells[0].merge(cells[2])
        cells[0].text = "No attachments recorded."
    style_docx_table(table, header=True, font_size=7)


def generate_observation_docx(path, record_ids=None):
    columns, data, rows = observation_export_rows(record_ids)
    if not rows:
        raise ValueError("There is no observation data to export.")
    doc = Document()
    set_docx_landscape(doc)
    # Keep the report compact and professional: one table, one row per current observation.
    add_docx_report_header(doc,"HSE OBSERVATION REPORT","OBS",f"{document_prefix()}-OBS-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}")
    pg=doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    pg.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    pg.paragraph_format.space_after=Pt(2)

    preferred = observation_word_sequence()
    idx = {name: columns.index(name) for name in preferred if name in columns}
    table = doc.add_table(rows=1, cols=len(preferred))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for c, name in enumerate(preferred):
        table.rows[0].cells[c].text = name

    for values, row in zip(data[:1000], rows[:1000]):
        cells = table.add_row().cells
        attachments = attachment_rows("observation_attachments", row["id"])
        for c, name in enumerate(preferred):
            cell=cells[c]
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if name == "Evidence":
                cell.text=""
                if not attachments:
                    cell.paragraphs[0].add_run("None").font.size=Pt(5.2)
                else:
                    for ai,a in enumerate(attachments):
                        p2=cell.paragraphs[0] if ai==0 else cell.add_paragraph()
                        p2.paragraph_format.space_after=Pt(0)
                        p2.paragraph_format.space_before=Pt(0)
                        path2=Path(safe(a["file_path"]))
                        if path2.exists() and path2.suffix.lower() in {".png",".jpg",".jpeg",".bmp"}:
                            try:
                                p2.add_run().add_picture(str(path2), width=Inches(0.38))
                            except Exception:
                                pass
                        run=p2.add_run(" " + path2.name)
                        run.font.size=Pt(4.8)
            elif name == "Observation Type":
                cell.text=xml_safe(row["obs_type"])
            elif name == "Observation Detail":
                cell.text=xml_safe(row["observation"])
            elif name == "Action Taken":
                cell.text=xml_safe(row["corrective_action"])
            elif name == "Responsible":
                cell.text=xml_safe(row["responsible"])
            elif name == "Category":
                cell.text=xml_safe(row["category"])
            elif name == "S.No.":
                cell.text=str(values[idx[name]])
            else:
                cell.text=xml_safe(values[idx[name]])

    # A4 landscape: total 10.95 inches, safely inside the 11.19-inch usable width.
    widths=[0.34,0.70,0.92,0.88,0.95,1.92,1.45,0.90,2.10,0.79]
    for row_i,rowx in enumerate(table.rows):
        trPr=rowx._tr.get_or_add_trPr()
        cant_split=OxmlElement("w:cantSplit"); trPr.append(cant_split)
        _set_docx_row_height(rowx, 2160, "exact")  # 1.5 inch for every observation-table row.
        for i,cell in enumerate(rowx.cells):
            cell.width=Inches(widths[i])
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p2 in cell.paragraphs:
                p2.paragraph_format.space_after=Pt(0)
                p2.paragraph_format.space_before=Pt(0)
                p2.paragraph_format.line_spacing=0.85
                for run in p2.runs:
                    run.font.size=Pt(5.2 if row_i else 5.3)
    style_docx_table(table, header=True, font_size=5.2)
    add_docx_footer(doc)
    save_docx_validated(doc, path)


def _set_docx_cell_margins(cell, top=40, start=50, bottom=40, end=50):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def _set_docx_row_height(row, height_twips, rule="atLeast"):
    trPr = row._tr.get_or_add_trPr()
    trHeight = trPr.find(qn("w:trHeight"))
    if trHeight is None:
        trHeight = OxmlElement("w:trHeight")
        trPr.append(trHeight)
    trHeight.set(qn("w:val"), str(height_twips))
    trHeight.set(qn("w:hRule"), rule)


def _card_cell_text(cell, label, value, label_width=None, font_size=6.4):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 0.85
    r = p.add_run(xml_safe(label))
    r.bold = True
    r.font.size = Pt(font_size)
    if value:
        r2 = p.add_run("  " + xml_safe(value))
        r2.font.size = Pt(font_size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _set_docx_cell_margins(cell)


def _add_card_image_evidence(cell, attachments):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    image_count = 0
    non_images = []
    for a in attachments:
        fp = Path(safe(a["file_path"]))
        if not fp.exists():
            continue
        if fp.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp"}:
            try:
                if image_count:
                    p.add_run("   ")
                run = p.add_run()
                run.add_picture(str(fp), width=Inches(0.72))
                image_count += 1
            except Exception:
                non_images.append(fp.name)
        else:
            non_images.append(fp.name)
    if image_count == 0 and not non_images:
        p.add_run("No evidence attached.").font.size = Pt(6)
    elif non_images:
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.space_before = Pt(0)
        r = p2.add_run("Files: " + "; ".join(non_images))
        r.font.size = Pt(5.6)
    return image_count


def generate_observation_card_docx(observation_ids, path):
    """Create one professional STOP Card per selected observation.
    Each card occupies exactly one A4 portrait page. Multiple selected observations
    are placed in the same Word file, one card per page.
    """
    if isinstance(observation_ids, (int, str)):
        observation_ids = [int(observation_ids)]
    ids = []
    for value in observation_ids or []:
        try:
            iv = int(value)
            if iv not in ids:
                ids.append(iv)
        except Exception:
            continue
    if not ids:
        raise ValueError("No observation was selected.")

    rows = []
    for oid in ids:
        row = db.fetchone("SELECT * FROM observations WHERE id=?", (oid,))
        if row:
            rows.append(row)
    if not rows:
        raise ValueError("The selected observation(s) no longer exist.")

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.left_margin = Inches(0.12)
    section.right_margin = Inches(0.12)
    section.top_margin = Inches(0.12)
    section.bottom_margin = Inches(0.12)
    section.header_distance = Inches(0.03)
    section.footer_distance = Inches(0.03)

    def add_border(cell, color="777777", sz=5):
        set_docx_cell_borders(cell,
            top={"val":"single","sz":sz,"color":color},
            bottom={"val":"single","sz":sz,"color":color},
            left={"val":"single","sz":sz,"color":color},
            right={"val":"single","sz":sz,"color":color})

    def add_text(cell, text, size=7.0, bold=False, align=None):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 0.82
        if align is not None:
            p.alignment = align
        r = p.add_run(xml_safe(text) if text else "")
        r.font.size = Pt(size)
        r.bold = bold
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _set_docx_cell_margins(cell, 18, 25, 18, 25)

    def add_section_heading(cell, title, size=7.0):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(xml_safe(title))
        r.bold = True
        r.font.size = Pt(size)
        add_docx_cell_shading(cell, "E7EDF7")
        _set_docx_cell_margins(cell, 18, 25, 18, 25)
        add_border(cell)

    for card_index, row in enumerate(rows):
        if card_index:
            doc.add_page_break()

        # Header: use the complete A4 width rather than the reduced sample size.
        header = doc.add_table(rows=1, cols=3)
        header.autofit = False
        header.alignment = WD_TABLE_ALIGNMENT.CENTER
        hw = [1.05, 5.70, 1.28]
        for i, w in enumerate(hw):
            header.rows[0].cells[i].width = Inches(w)
            _set_docx_cell_margins(header.rows[0].cells[i], 20, 30, 20, 30)
            add_border(header.rows[0].cells[i], sz=6)

        c = header.rows[0].cells[0]
        c.text = ""
        logo_path = report_logo_path()
        if logo_path and Path(logo_path).exists():
            try:
                c.paragraphs[0].add_run().add_picture(logo_path, width=Inches(0.82))
            except Exception:
                add_text(c, company_name(), 7.5, True, WD_ALIGN_PARAGRAPH.CENTER)
        else:
            add_text(c, company_name(), 7.5, True, WD_ALIGN_PARAGRAPH.CENTER)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        c = header.rows[0].cells[1]
        c.text = ""
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("HSE OBSERVATION AND STOP CARD")
        r.bold = True; r.font.size = Pt(13)

        c = header.rows[0].cells[2]
        c.text = ""
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("Document Number\n"); r.bold = True; r.font.size = Pt(6.5)
        r = p.add_run(xml_safe(row["number"])); r.bold = True; r.font.size = Pt(8)

        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(1)
        spacer.paragraph_format.space_before = Pt(0)

        outer = doc.add_table(rows=1, cols=2)
        outer.autofit = False
        outer.alignment = WD_TABLE_ALIGNMENT.CENTER
        outer_widths = [3.95, 3.96]
        for i, w in enumerate(outer_widths):
            outer.rows[0].cells[i].width = Inches(w)
            outer.rows[0].cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            _set_docx_cell_margins(outer.rows[0].cells[i], 12, 14, 12, 14)
            add_border(outer.rows[0].cells[i], sz=6)

        # LEFT SIDE ---------------------------------------------------------
        left = outer.rows[0].cells[0]
        left.text = ""
        lt = left.add_table(rows=1, cols=2)
        lt.autofit = False
        lt.alignment = WD_TABLE_ALIGNMENT.CENTER
        lt.columns[0].width = Inches(1.30)
        lt.columns[1].width = Inches(2.57)
        lt.rows[0].cells[0].merge(lt.rows[0].cells[1])
        add_section_heading(lt.rows[0].cells[0], "OBSERVATION / STOP CARD", 7.5)

        fields = [
            ("Originator's Name", row["observer"]),
            ("Company", row["observer_company"] or row["company"]),
            ("Observer Designation", row["observer_designation"]),
            ("Time and Date", f'{safe(row["obs_date"])}  {safe(row["obs_time"])}'.strip()),
            ("Location", row["location"]),
            ("Project / Area", " / ".join(x for x in [safe(row["project"]), safe(row["area"])] if x)),
            ("Reported / Responsible", row["responsible"]),
            ("Status", row["status"]),
            ("Target Date", row["target_date"]),
        ]
        for label, value in fields:
            cells = lt.add_row().cells
            cells[0].width = Inches(1.30); cells[1].width = Inches(2.57)
            add_text(cells[0], label, 6.1, True)
            add_text(cells[1], value, 6.5)
            add_docx_cell_shading(cells[0], "F0F0F0")
            add_border(cells[0]); add_border(cells[1])

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Observation Type", 7.0)
        p = cells[0].add_paragraph()
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 0.75
        for i, typ in enumerate(OBS_TYPES):
            mark = "☑" if typ == safe(row["obs_type"]) else "☐"
            rr = p.add_run(f"{mark} {typ}")
            rr.font.size = Pt(5.9)
            if i < len(OBS_TYPES)-1:
                rr.add_text("   ")

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Observation Details", 7.0)
        p = cells[0].add_paragraph(xml_safe(row["observation"]) or "No observation detail recorded.")
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0); p.paragraph_format.line_spacing = 0.82
        for rr in p.runs: rr.font.size = Pt(7.2)

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Action Taken / Corrective Action", 7.0)
        p = cells[0].add_paragraph(xml_safe(row["corrective_action"]) or "No action recorded.")
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0); p.paragraph_format.line_spacing = 0.82
        for rr in p.runs: rr.font.size = Pt(7.2)

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Related Factors / Category", 7.0)
        p = cells[0].add_paragraph()
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        related = safe(row["category"])
        if safe(row["subcategory"]):
            related += " / " + safe(row["subcategory"])
        rr = p.add_run(("☑ " + related) if related else "☐ Not specified")
        rr.font.size = Pt(6.7)

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Evidence", 7.0)
        _add_card_image_evidence(cells[0], attachment_rows("observation_attachments", row["id"]))

        # RIGHT SIDE --------------------------------------------------------
        right = outer.rows[0].cells[1]
        right.text = ""
        rt = right.add_table(rows=1, cols=2)
        rt.autofit = False
        rt.alignment = WD_TABLE_ALIGNMENT.CENTER
        rt.columns[0].width = Inches(1.12)
        rt.columns[1].width = Inches(2.72)
        rt.rows[0].cells[0].merge(rt.rows[0].cells[1])
        add_section_heading(rt.rows[0].cells[0], "HAZARD IDENTIFICATION CARD", 7.5)

        selected_cat = safe(row["category"])
        selected_subcat = safe(row["subcategory"])
        selected_priority = safe(row["priority"])
        selected_status = safe(row["status"])

        hazard_cells = rt.add_row().cells
        hazard_cells[0].width = Inches(1.12); hazard_cells[1].width = Inches(2.72)
        add_text(hazard_cells[0], "Hazard Category", 6.2, True, WD_ALIGN_PARAGRAPH.CENTER)
        add_docx_cell_shading(hazard_cells[0], "F0F0F0")
        hazard_cells[1].text = ""
        # Three compact columns: every menu category remains visible without using the sample card.
        grid = hazard_cells[1].add_table(rows=0, cols=3)
        grid.autofit = False
        for col in grid.columns:
            col.width = Inches(0.88)
        n = len(CATEGORIES)
        per_col = (n + 2) // 3
        for rix in range(per_col):
            cells3 = grid.add_row().cells
            for j in range(3):
                idx_cat = rix + j * per_col
                text_cat = CATEGORIES[idx_cat] if idx_cat < n else ""
                add_text(cells3[j], (("☑ " if text_cat == selected_cat else "☐ ") + text_cat) if text_cat else "", 5.0)
                if text_cat == selected_cat:
                    for run in cells3[j].paragraphs[0].runs:
                        run.bold = True
        if selected_subcat:
            p = hazard_cells[1].add_paragraph()
            p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
            rr = p.add_run("Subcategory: " + selected_subcat); rr.bold = True; rr.font.size = Pt(5.8)
        add_border(hazard_cells[0]); add_border(hazard_cells[1])

        def add_card_checklist(label, options, selected_value, font_size=5.9):
            cells = rt.add_row().cells
            add_text(cells[0], label, 6.2, True, WD_ALIGN_PARAGRAPH.CENTER)
            add_docx_cell_shading(cells[0], "F0F0F0")
            cells[1].text = ""
            p = cells[1].paragraphs[0]
            p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0); p.paragraph_format.line_spacing = 0.76
            for option in options:
                mark = "☑" if option == selected_value else "☐"
                rr = p.add_run(f"{mark} {option}   ")
                rr.font.size = Pt(font_size)
            add_border(cells[0]); add_border(cells[1])

        add_card_checklist("Priority", PRIORITIES, selected_priority, 6.0)
        add_card_checklist("Status", STATUSES, selected_status, 6.0)

        action_cells = rt.add_row().cells
        add_text(action_cells[0], "Action / Control", 6.2, True, WD_ALIGN_PARAGRAPH.CENTER)
        add_docx_cell_shading(action_cells[0], "F0F0F0")
        add_text(action_cells[1], row["corrective_action"] or "No action recorded.", 6.2)
        add_border(action_cells[0]); add_border(action_cells[1])

        # Keep the two nested tables visually compact and prevent row splitting.
        for nested in (lt, rt, grid):
            for rr in nested.rows:
                trPr = rr._tr.get_or_add_trPr()
                cant = OxmlElement("w:cantSplit")
                trPr.append(cant)
                for cell in rr.cells:
                    _set_docx_cell_margins(cell, 14, 20, 14, 20)

        # Keep the whole outer card on one page where possible.
        for rr in outer.rows:
            trPr = rr._tr.get_or_add_trPr()
            trPr.append(OxmlElement("w:cantSplit"))

    add_docx_footer(doc)
    save_docx_validated(doc, path)

def generate_table_docx(table_name, path, record_ids=None):
    if table_name == "observations":
        return generate_observation_docx(path, record_ids=record_ids)
    columns, data = MainWindow.table_data_static(table_name, record_ids=record_ids)
    if not columns:
        raise ValueError("There is no data to export.")
    doc = Document()
    add_docx_header(doc)
    add_docx_title(doc, f"{table_name.replace('_', ' ').title()} Report")
    p = doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    table = doc.add_table(rows=1, cols=len(columns))
    for i, col in enumerate(columns):
        table.rows[0].cells[i].text = xml_safe(col)
    if table_name == "incidents" and "Attachments" in columns:
        incident_rows = db.fetchall("SELECT * FROM incidents ORDER BY id DESC")
        if record_ids:
            wanted=set(record_ids); incident_rows=[r for r in incident_rows if r["id"] in wanted]
        attach_col=columns.index("Attachments")
        for r in incident_rows[:1000]:
            cells=table.add_row().cells
            # Populate the same register fields as table_data_static, but replace the
            # attachment-count/text cell with the actual evidence files.
            for i,col in enumerate(columns):
                if col != "Attachments":
                    cells[i].text=xml_safe(r[col] if col in r.keys() else "")
            attachments=attachment_rows("incident_attachments",r["id"])
            if not attachments:
                cells[attach_col].text="No attachment"
            else:
                cells[attach_col].text=""
                for a in attachments:
                    fp=Path(safe(a["file_path"]))
                    remark=safe(a["remark"]) if "remark" in a.keys() else ""
                    if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                        try:
                            pp=cells[attach_col].paragraphs[0] if not cells[attach_col].paragraphs[0].runs else cells[attach_col].add_paragraph()
                            pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
                            pp.add_run().add_picture(str(fp),width=Inches(1.25))
                            cap=cells[attach_col].add_paragraph(fp.name + (f" | {remark}" if remark else ""))
                            cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
                        except Exception:
                            cells[attach_col].add_paragraph(fp.name + (f" | {remark}" if remark else ""))
                    else:
                        cells[attach_col].add_paragraph(fp.name + (f" | {remark}" if remark else ""))
    else:
        for row in data[:1000]:
            cells = table.add_row().cells
            for i, value in enumerate(row):
                cells[i].text = xml_safe(value)
    style_docx_table(table)
    add_docx_footer(doc)
    save_docx_validated(doc, path)


class PieChartWidget(QWidget):
    def __init__(self, values=None, parent=None):
        super().__init__(parent)
        self.values = values or {}
        self.setMinimumHeight(260)

    def set_values(self, values):
        self.values = values or {}
        self.update()

    def paintEvent(self, event):
        painter=QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect=self.rect().adjusted(20,20,-20,-20)
        total=sum(max(0,int(v)) for v in self.values.values())
        if total<=0:
            painter.drawText(rect,Qt.AlignmentFlag.AlignCenter,"No observation data")
            painter.end(); return
        center=rect.center()
        radius=min(rect.width(),rect.height())//3
        start=0
        palette=[Qt.GlobalColor.darkBlue,Qt.GlobalColor.darkGreen,Qt.GlobalColor.darkRed,
                 Qt.GlobalColor.darkCyan,Qt.GlobalColor.darkMagenta,Qt.GlobalColor.darkYellow,
                 Qt.GlobalColor.gray]
        for i,(label,value) in enumerate(self.values.items()):
            value=max(0,int(value))
            span=int(round(360*16*value/total))
            painter.setBrush(QBrush(palette[i%len(palette)]))
            painter.setPen(QPen(Qt.GlobalColor.white,1))
            painter.drawPie(QRectF(center.x()-radius,center.y()-radius,2*radius,2*radius),start,span)
            start += span
        y=25
        x=rect.left()
        painter.setFont(QFont("Arial",9))
        for i,(label,value) in enumerate(self.values.items()):
            painter.setBrush(QBrush(palette[i%len(palette)]))
            painter.drawRect(x,y,14,14)
            painter.setPen(QPen(Qt.GlobalColor.black))
            painter.drawText(x+20,y+12,f"{label}: {value}")
            y += 20
        painter.end()


class BarChartWidget(QWidget):
    def __init__(self, values=None, parent=None):
        super().__init__(parent)
        self.values=values or {}
        self.setMinimumHeight(260)

    def set_values(self, values):
        self.values=values or {}
        self.update()

    def paintEvent(self,event):
        painter=QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect=self.rect().adjusted(45,20,-20,-45)
        if not self.values:
            painter.drawText(rect,Qt.AlignmentFlag.AlignCenter,"No data")
            painter.end(); return
        maxv=max([int(v) for v in self.values.values()] or [1])
        count=len(self.values)
        barw=max(18,int(rect.width()/max(count,1)*0.65))
        gap=max(8,int(rect.width()/max(count,1)*0.35))
        x=rect.left()
        colors=[Qt.GlobalColor.darkBlue,Qt.GlobalColor.darkGreen,Qt.GlobalColor.darkRed,
                Qt.GlobalColor.darkCyan,Qt.GlobalColor.darkMagenta,Qt.GlobalColor.darkYellow]
        for i,(label,value) in enumerate(self.values.items()):
            value=int(value)
            h=int(rect.height()*value/maxv) if maxv else 0
            y=rect.bottom()-h
            painter.setBrush(QBrush(colors[i%len(colors)]))
            painter.setPen(QPen(Qt.GlobalColor.black))
            painter.drawRect(x,y,barw,h)
            painter.setPen(QPen(Qt.GlobalColor.black))
            painter.drawText(x,y-5,str(value))
            painter.save()
            painter.translate(x+barw/2,rect.bottom()+12)
            painter.rotate(-45)
            painter.drawText(0,0,str(label)[:16])
            painter.restore()
            x += barw+gap
        painter.drawLine(rect.left(),rect.bottom(),rect.right(),rect.bottom())
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
        w,layout=self.page("HSE Dashboard")
        cards=QHBoxLayout()
        values=[
            ("Observations","SELECT COUNT(*) c FROM observations"),
            ("Open Observations","SELECT COUNT(*) c FROM observations WHERE status NOT IN ('Closed','Cancelled')"),
            ("Overdue","SELECT COUNT(*) c FROM observations WHERE status NOT IN ('Closed','Cancelled') AND target_date < date('now')"),
            ("Incidents","SELECT COUNT(*) c FROM incidents"),
            ("Audits","SELECT COUNT(*) c FROM audits"),
            ("CAPA","SELECT COUNT(*) c FROM capa")
        ]
        for name,sql in values:
            value=db.fetchone(sql)["c"]
            box=QGroupBox(name); bl=QVBoxLayout(box)
            label=QLabel(str(value)); label.setStyleSheet("font-size:28px;font-weight:bold;color:#17365D;")
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            bl.addWidget(label); cards.addWidget(box)
        layout.addLayout(cards)
        layout.addWidget(QLabel(f"<b>Company:</b> {company_name()}    <b>Project:</b> {db.setting('project_name','')}"))
        selector=QComboBox(); selector.addItems(["HSE Home Dashboard","Inspection Dashboard","Incident Dashboard"]); layout.addWidget(selector)
        def switch_dashboard(text):
            if text=="Inspection Dashboard": self.observations()
            elif text=="Incident Dashboard": self.incidents()
        selector.currentTextChanged.connect(switch_dashboard)

        charts=QHBoxLayout()
        pie=PieChartWidget()
        obs_types=db.fetchall("SELECT obs_type,COUNT(*) c FROM observations GROUP BY obs_type ORDER BY obs_type COLLATE NOCASE")
        pie.set_values({safe(r["obs_type"]) or "Unspecified":r["c"] for r in obs_types})
        pie_box=QGroupBox("Observation Type Distribution (no categories merged)"); pl=QVBoxLayout(pie_box); pl.addWidget(pie)
        bar=BarChartWidget()
        stats={"Observations":db.fetchone("SELECT COUNT(*) c FROM observations")["c"],
               "Incidents":db.fetchone("SELECT COUNT(*) c FROM incidents")["c"],
               "Audits":db.fetchone("SELECT COUNT(*) c FROM audits")["c"],
               "CAPA":db.fetchone("SELECT COUNT(*) c FROM capa")["c"]}
        bar.set_values(stats)
        bar_box=QGroupBox("HSE Register Summary"); bl=QVBoxLayout(bar_box); bl.addWidget(bar)
        charts.addWidget(pie_box,1); charts.addWidget(bar_box,1); layout.addLayout(charts)

        # A clear, complete category register. Every category is a separate row, including zero counts.
        cat_box=QGroupBox("HSE Observation Categories — click a row to view that category")
        cat_layout=QVBoxLayout(cat_box)
        cat_table=QTableWidget()
        categories=sorted(set(CATEGORIES), key=lambda x: x.lower())
        counts={safe(r["category"]):int(r["c"]) for r in db.fetchall("SELECT category,COUNT(*) c FROM observations GROUP BY category")}
        cat_table.setColumnCount(3)
        cat_table.setHorizontalHeaderLabels(["Category","Observations","Open / Active"])
        cat_table.setRowCount(len(categories))
        cat_table.horizontalHeader().setSectionResizeMode(0,QHeaderView.Stretch)
        cat_table.horizontalHeader().setSectionResizeMode(1,QHeaderView.ResizeToContents)
        cat_table.horizontalHeader().setSectionResizeMode(2,QHeaderView.ResizeToContents)
        cat_table.setAlternatingRowColors(True)
        cat_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        cat_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        for r,cat in enumerate(categories):
            total=counts.get(cat,0)
            active=db.fetchone("SELECT COUNT(*) c FROM observations WHERE category=? AND status NOT IN ('Closed','Cancelled')",(cat,))["c"]
            cat_table.setItem(r,0,QTableWidgetItem(cat))
            cat_table.setItem(r,1,QTableWidgetItem(str(total)))
            cat_table.setItem(r,2,QTableWidgetItem(str(active)))
        cat_table.setMinimumHeight(min(330, max(150, len(categories)*26+35)))
        def open_category(row,_column):
            item=cat_table.item(row,0)
            if item:
                self.observation_dashboard_filter=item.text()
                self.observations()
        cat_table.cellDoubleClicked.connect(open_category)
        cat_layout.addWidget(cat_table)
        layout.addWidget(cat_box,1)

        recent_inc=QGroupBox("Recent Incidents"); ril=QVBoxLayout(recent_inc); it=QTableWidget(); it.setColumnCount(5); it.setHorizontalHeaderLabels(["ID / Number","Date","Type","Status","Completion"]); ir=db.fetchall("SELECT id,number,incident_date,incident_type,status,incident_title,description,investigation_method,direct_cause,root_cause FROM incidents ORDER BY id DESC LIMIT 10"); it.setRowCount(len(ir))
        for rr,x in enumerate(ir):
            total=10; done=sum(bool(safe(x[k])) for k in ["incident_date","incident_time"] if k in x.keys())
            done += sum(bool(safe(x[k])) for k in ["location","project","incident_type","description","investigation_method","direct_cause","root_cause"] if k in x.keys())
            pct=round(done*100/total)
            for cc,v in enumerate([f"{safe(x['id'])} / {safe(x['number'])}",x['incident_date'],x['incident_type'],x['status'],f"{pct}%"]): it.setItem(rr,cc,QTableWidgetItem(safe(v)))
        ril.addWidget(it); layout.addWidget(recent_inc)

        recent=QGroupBox("Recent Observations")
        recent_layout=QVBoxLayout(recent)
        table=QTableWidget(); table.setColumnCount(8)
        table.setHorizontalHeaderLabels(["S.No.","Observation No.","Date","Type","Category","Location","Status","Action Taken"])
        table.horizontalHeader().setSectionResizeMode(7,QHeaderView.Stretch)
        for c in range(7): table.horizontalHeader().setSectionResizeMode(c,QHeaderView.ResizeToContents)
        rows=db.fetchall("SELECT id,number,obs_date,obs_type,category,location,status,corrective_action FROM observations ORDER BY id DESC LIMIT 10")
        table.setRowCount(len(rows))
        for r,row in enumerate(rows):
            vals=[str(r+1),row["number"],row["obs_date"],row["obs_type"],row["category"],row["location"],row["status"],row["corrective_action"]]
            for c,v in enumerate(vals):
                item=QTableWidgetItem(safe(v)); item.setToolTip(safe(v)); table.setItem(r,c,item)
        table.setAlternatingRowColors(True); table.setWordWrap(True)
        recent_layout.addWidget(table)
        layout.addWidget(recent,1)


    def observations(self):
        w, layout = self.page("HSE Inspection Register")
        toolbar = QHBoxLayout()
        search = QLineEdit()
        search.setPlaceholderText("Search observation no., location, responsible person, description, action...")
        toolbar.addWidget(search)
        clear_filter=QPushButton("Clear Category Filter")
        toolbar.addWidget(clear_filter)
        add_button = QPushButton("+ New Observation")
        attachment_button = QPushButton("Attachments")
        delete_button = QPushButton("Delete Selected")
        export_button = QPushButton("Export CSV")
        excel_button = QPushButton("Export Excel")
        pdf_button = QPushButton("Export PDF")
        word_button = QPushButton("Export Word")
        card_button = QPushButton("Draft STOP Card(s)")
        for b in [add_button, attachment_button, delete_button, export_button, excel_button, pdf_button, word_button, card_button]:
            toolbar.addWidget(b)
        layout.addLayout(toolbar)

        table = QTableWidget()
        headers = ["ID", "S.No.", "Observation No.", "Observation Date", "Observation Type", "Observation Detail", "Status",
                   "Location", "Action Required / Corrective Action", "Target Date", "Attachments"]
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setColumnHidden(0, True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        table.setWordWrap(True)
        table.setAlternatingRowColors(True)
        layout.addWidget(table)

        def load():
            term = search.text().strip()
            category_filter = getattr(self, "observation_dashboard_filter", "")
            query = "SELECT * FROM observations"
            clauses=[]; params=[]
            if term:
                clauses.append("(number LIKE ? OR location LIKE ? OR responsible LIKE ? OR observation LIKE ? OR corrective_action LIKE ? OR category LIKE ?)")
                like = f"%{term}%"
                params.extend([like, like, like, like, like, like])
            if category_filter:
                clauses.append("category=?"); params.append(category_filter)
            if clauses:
                query += " WHERE " + " AND ".join(clauses)
            query += " ORDER BY id DESC"
            rows = db.fetchall(query, tuple(params))
            table.setRowCount(len(rows))
            for r, row in enumerate(rows):
                attachments = attachment_names("observation_attachments", row["id"])
                values = [
                    row["id"], r + 1, row["number"], row["obs_date"], row["obs_type"], row["observation"], row["status"],
                    row["location"], row["corrective_action"], row["target_date"], attachments or "No attachment"
                ]
                for c, value in enumerate(values):
                    item = QTableWidgetItem(safe(value))
                    item.setToolTip(safe(value))
                    table.setItem(r, c, item)
                if overdue(row["target_date"], row["status"]):
                    table.item(r, 9).setForeground(Qt.GlobalColor.red)
                    table.item(r, 6).setForeground(Qt.GlobalColor.red)
                table.setRowHeight(r, 48)

        def selected_observation_ids(require_one=False):
            indexes = table.selectionModel().selectedRows(0) if table.selectionModel() else []
            ids = []
            for index in indexes:
                item = table.item(index.row(), 0)
                if item:
                    try:
                        oid = int(item.text())
                        if oid not in ids:
                            ids.append(oid)
                    except Exception:
                        pass
            if not ids:
                current = table.currentRow()
                if current >= 0 and table.item(current, 0):
                    try:
                        ids = [int(table.item(current, 0).text())]
                    except Exception:
                        ids = []
            if require_one and len(ids) != 1:
                QMessageBox.warning(self, "Observation", "Select exactly one observation for this action.")
                return []
            if not ids:
                QMessageBox.warning(self, "Observation", "Select at least one observation first.")
            return ids

        def selected_observation_id():
            ids = selected_observation_ids(require_one=True)
            return ids[0] if ids else None

        def draft_selected_card():
            obs_ids = selected_observation_ids()
            if not obs_ids:
                return
            records = []
            for obs_id in obs_ids:
                record = db.fetchone("SELECT number FROM observations WHERE id=?", (obs_id,))
                if record:
                    records.append(record)
            if not records:
                QMessageBox.warning(self, "Observation", "The selected observation(s) could not be found.")
                return
            if len(records) == 1:
                default_name = f"{safe(records[0]['number'])}_Observation_STOP_Card.docx"
            else:
                default_name = "Selected_Observations_STOP_Cards.docx"
            path, _ = QFileDialog.getSaveFileName(
                self, "Draft Observation STOP Card(s)", default_name, "Word Documents (*.docx)"
            )
            if not path:
                return
            try:
                generate_observation_card_docx(obs_ids, path)
                self.show_export_success(path)
            except Exception as e:
                logging.exception("Observation card report failed")
                QMessageBox.critical(self, "Report Error", str(e))

        def manage_attachments():
            obs_id = selected_observation_id()
            if obs_id is None:
                return
            record = db.fetchone("SELECT number FROM observations WHERE id=?", (obs_id,))
            if not record:
                return
            dialog = QDialog(self)
            dialog.setWindowTitle(f"Attachments - {record['number']}")
            dialog.resize(760, 480)
            v = QVBoxLayout(dialog)
            v.addWidget(QLabel(f"<b>Observation:</b> {safe(record['number'])}"))
            listw = QListWidget()
            v.addWidget(listw, 1)
            def refresh_files():
                listw.clear()
                rows = attachment_rows("observation_attachments", obs_id)
                for a in rows:
                    p = Path(safe(a["file_path"]))
                    listw.addItem(f"{p.name}  |  {safe(a['attachment_type'])}")
                if not rows:
                    listw.addItem("No attachments recorded.")
            def add_files():
                paths, _ = QFileDialog.getOpenFileNames(
                    dialog, "Add Observation Evidence", "",
                    "Evidence Files (*.png *.jpg *.jpeg *.bmp *.pdf *.doc *.docx *.xls *.xlsx *.txt *.mp4 *.avi *.mov);;All Files (*)"
                )
                if paths:
                    copy_attachments(paths, safe(record["number"]), "observation_attachments", obs_id)
                    refresh_files(); load()
            add = QPushButton("Add Attachment")
            close = QPushButton("Close")
            rowbtn = QHBoxLayout(); rowbtn.addWidget(add); rowbtn.addStretch(); rowbtn.addWidget(close); v.addLayout(rowbtn)
            add.clicked.connect(add_files); close.clicked.connect(dialog.accept)
            refresh_files(); dialog.exec()

        def delete_selected():
            obs_id = selected_observation_id()
            if obs_id is None:
                return
            record = db.fetchone("SELECT number FROM observations WHERE id=?", (obs_id,))
            if not record:
                return
            answer = QMessageBox.question(
                self, "Confirm Delete",
                f"Delete observation {record['number']} and its evidence?\n\nThis cannot be undone."
            )
            if answer != QMessageBox.Yes:
                return
            for a in attachment_rows("observation_attachments", obs_id):
                try:
                    Path(safe(a["file_path"])).unlink(missing_ok=True)
                except Exception:
                    logging.exception("Unable to remove observation attachment")
            db.execute("DELETE FROM observations WHERE id=?", (obs_id,))
            load()

        def clear_category_filter():
            self.observation_dashboard_filter=""
            search.clear()
            load()
        clear_filter.clicked.connect(clear_category_filter)
        search.textChanged.connect(load)
        add_button.clicked.connect(lambda: self.observation_form(load))
        attachment_button.clicked.connect(manage_attachments)
        delete_button.clicked.connect(delete_selected)
        export_button.clicked.connect(lambda: self.export_csv("observations"))
        excel_button.clicked.connect(lambda: self.export_excel("observations"))
        pdf_button.clicked.connect(lambda: self.export_pdf("observations"))
        word_button.clicked.connect(lambda: self.export_docx("observations"))
        card_button.clicked.connect(draft_selected_card)
        load()


    def observation_form(self, refresh):
        dialog = QDialog(self)
        dialog.setWindowTitle("New HSE Observation / Inspection")
        dialog.resize(980, 760)
        dialog.setMinimumSize(900, 650)
        outer = QVBoxLayout(dialog)

        header = QLabel("HSE Inspection Observation")
        header.setStyleSheet("font-size:20px;font-weight:bold;color:#17365D;padding:6px;")
        outer.addWidget(header)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        content = QWidget()
        form = QFormLayout(content)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        edits = {}

        def edit(name):
            e = QLineEdit()
            edits[name] = e
            form.addRow(name.replace("_", " ").title() + ":", e)
            return e

        observation_date = QDateEdit(); observation_date.setCalendarPopup(True); observation_date.setDate(datetime.now().date())
        form.addRow("Observation Date:", observation_date)

        for name in ["project", "company", "location", "area", "responsible", "designation", "employee_id",
                     "observer", "observer_designation", "observer_id", "observer_company"]:
            edit(name)

        obs_type = QComboBox(); obs_type.addItems(OBS_TYPES)
        form.addRow("Observation Type:", obs_type)
        category = QComboBox(); category.addItems(CATEGORIES)
        form.addRow("Category:", category)
        edit("subcategory")

        observation = QTextEdit(); observation.setMinimumHeight(100)
        form.addRow("Observation Description:", observation)
        corrective = QTextEdit(); corrective.setMinimumHeight(90)
        form.addRow("Action Required / Corrective Action:", corrective)

        priority = QComboBox(); priority.addItems(PRIORITIES)
        form.addRow("Priority:", priority)
        target = QDateEdit(); target.setCalendarPopup(True); target.setDate(datetime.now().date())
        form.addRow("Target Completion Date:", target)
        status = QComboBox(); status.addItems(STATUSES)
        form.addRow("Status:", status)

        attachment_paths = []
        attachment_label = QLabel("No evidence files selected.")
        attachment_label.setWordWrap(True)
        attach_button = QPushButton("Attach Evidence Files")
        attach_button.setMinimumHeight(36)

        def choose_files():
            paths, _ = QFileDialog.getOpenFileNames(
                dialog, "Select Evidence Files", "",
                "Evidence Files (*.png *.jpg *.jpeg *.bmp *.pdf *.doc *.docx *.xls *.xlsx *.txt *.mp4 *.avi *.mov);;All Files (*)"
            )
            if paths:
                attachment_paths.clear()
                attachment_paths.extend(paths)
                attachment_label.setText("\n".join(Path(p).name for p in paths))

        attach_button.clicked.connect(choose_files)
        form.addRow("Evidence / Attachments:", attach_button)
        form.addRow("Selected Files:", attachment_label)

        scroll_area.setWidget(content)
        outer.addWidget(scroll_area, 1)

        buttons = QDialogButtonBox()
        save_button = buttons.addButton("Save Observation", QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
        save_button.setMinimumHeight(40)
        cancel_button.setMinimumHeight(40)
        outer.addWidget(buttons)
        cancel_button.clicked.connect(dialog.reject)

        def save():
            if not edits["location"].text().strip():
                QMessageBox.warning(dialog, "Required", "Location is required.")
                return
            if not observation.toPlainText().strip():
                QMessageBox.warning(dialog, "Required", "Observation description is required.")
                return
            try:
                number = next_number("HSE-OBS", "observations")
                db.execute("""
                    INSERT INTO observations (
                        number, obs_date, obs_time, project, company, location, area,
                        responsible, designation, employee_id, observer, observer_designation,
                        observer_id, observer_company, obs_type, category, subcategory,
                        observation, corrective_action, priority, target_date, status, created_at
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    number, observation_date.date().toString("yyyy-MM-dd"), now_time(), edits["project"].text(), edits["company"].text(),
                    edits["location"].text(), edits["area"].text(), edits["responsible"].text(),
                    edits["designation"].text(), edits["employee_id"].text(), edits["observer"].text(),
                    edits["observer_designation"].text(), edits["observer_id"].text(),
                    edits["observer_company"].text(), obs_type.currentText(), category.currentText(),
                    edits["subcategory"].text(), observation.toPlainText(), corrective.toPlainText(),
                    priority.currentText(), target.date().toString("yyyy-MM-dd"), status.currentText(), datetime.now().isoformat()
                ))
                row = db.fetchone("SELECT id FROM observations WHERE number=?", (number,))
                if row and attachment_paths:
                    copy_attachments(attachment_paths, number, "observation_attachments", row["id"])
                dialog.accept()
                refresh()
            except Exception as e:
                logging.exception("Observation save failed")
                QMessageBox.critical(dialog, "Save Error", f"Unable to save observation.\n\n{e}")

        save_button.clicked.connect(save)
        dialog.exec()


    def incidents(self):
        w, layout = self.page("Accident / Incident Investigation")
        banner=QHBoxLayout(); logo=QLabel(); logo.setFixedSize(80,60); logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_path=db.setting("company_logo","")
        if logo_path and Path(logo_path).exists():
            logo.setPixmap(QPixmap(logo_path).scaled(70,55,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))
        banner.addWidget(logo); banner.addWidget(QLabel(f"<b>{safe(company_name())}</b><br>Accident / Incident Investigation Register"),1); layout.addLayout(banner)
        toolbar=QHBoxLayout()
        buttons=[("+ New Incident", None),("Edit Selected",None),("Delete Selected",None),("Export CSV",lambda:self.export_csv("incidents")),( "Export Excel",lambda:self.export_excel("incidents")),( "Export PDF",lambda:self.export_pdf("incidents")),( "Export Word",lambda:self.export_docx("incidents")),( "Professional Report",None)]
        btns={}
        for text,fn in buttons:
            b=QPushButton(text); toolbar.addWidget(b); btns[text]=b
            if fn:b.clicked.connect(fn)
        layout.addLayout(toolbar)
        table=QTableWidget(); headers=["ID","Number","Date","Type","Location","Project","Method","Status","Completion","Attachments","Professional Report"]
        table.setColumnCount(len(headers)); table.setHorizontalHeaderLabels(headers); table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents); table.horizontalHeader().setStretchLastSection(True); table.setSelectionBehavior(QAbstractItemView.SelectRows); layout.addWidget(table)
        def load():
            rows=db.fetchall("SELECT * FROM incidents ORDER BY id DESC"); table.setRowCount(len(rows))
            for r,row in enumerate(rows):
                # Show the actual evidence file names in the register instead of an
                # attachment count, so the user can immediately see what is attached.
                attached_names=attachment_names("incident_attachments",row["id"]) or "No attachment"
                
                total=10; done=sum(bool(safe(row[k])) for k in ["incident_date","incident_time","location","project","incident_type","description","immediate_action","investigation_method","direct_cause","root_cause"])
                pct=round(done*100/total)
                try:
                    report_details=json.loads(safe(row["investigation_details"]) or "{}")
                except Exception:
                    report_details={}
                vals=[row["id"],row["number"],row["incident_date"],row["incident_type"],row["location"],row["project"],row["investigation_method"],row["status"],f"{pct}%",attached_names,"Available" if report_details.get("professional_report_generated") else "Not Generated"]
                for c,v in enumerate(vals):table.setItem(r,c,QTableWidgetItem(safe(v)))
        def professional():
            incident_id=self.selected_id(table,"Professional Report")
            if incident_id is None:return
            number=table.item(table.currentRow(),1).text(); path,_=QFileDialog.getSaveFileName(self,"Save Investigation Report",f"{number}_Investigation_Report.docx","Word Document (*.docx)")
            if not path:return
            try:
                generate_incident_docx(incident_id,path)
                rr=db.fetchone("SELECT investigation_details FROM incidents WHERE id=?",(incident_id,))
                try: details=json.loads(safe(rr["investigation_details"]) or "{}") if rr else {}
                except Exception: details={}
                details["professional_report_generated"]=True; details["professional_report_path"]=str(path)
                db.execute("UPDATE incidents SET investigation_details=? WHERE id=?",(json.dumps(details,ensure_ascii=False),incident_id))
                self.show_export_success(path); load()
            except Exception as e: logging.exception("Incident Word report failed"); QMessageBox.critical(self,"Report Error",str(e))
        btns["+ New Incident"].clicked.connect(lambda:self.incident_form(load))
        def edit_selected():
            rid=self.selected_id(table,"Edit Incident")
            if rid is None:return
            rr=db.fetchone("SELECT status FROM incidents WHERE id=?",(rid,))
            if rr and safe(rr["status"]).lower()=="closed": QMessageBox.information(self,"Submitted","This report has been submitted and cannot be edited."); return
            self.incident_form(load,rid)
        btns["Edit Selected"].clicked.connect(edit_selected)
        btns["Delete Selected"].clicked.connect(lambda:self.delete_selected_record(table,"incidents","incident_attachments","Incident",load))
        btns["Professional Report"].clicked.connect(professional)
        load()


    def incident_form(self, refresh, incident_id=None):
        existing = db.fetchone("SELECT * FROM incidents WHERE id=?", (incident_id,)) if incident_id else None
        if incident_id and not existing: return
        dialog=QDialog(self); dialog.setWindowTitle("Edit Accident / Incident Investigation" if existing else "New Accident / Incident Investigation"); dialog.resize(1200,850); dialog.setMinimumSize(1100,750)
        outer=QVBoxLayout(dialog); title=QLabel("ACCIDENT / INCIDENT INVESTIGATION"); title.setStyleSheet("font-size:20px;font-weight:bold;color:#17365D;padding:6px;"); outer.addWidget(title)
        scroll=QScrollArea(); scroll.setWidgetResizable(True); content=QWidget(); form=QFormLayout(content); form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow); edits={}
        def add_edit(name,label=None,value=""):
            e=QLineEdit(safe(value)); edits[name]=e; form.addRow(label or name.replace("_"," ").title()+":",e); return e
        add_edit("incident_title","Incident Title",existing["incident_title"] if existing and "incident_title" in existing.keys() else "")
        for name in ["incident_time","location","project","company","department","activity"]: add_edit(name,value=existing[name] if existing else "")
        for name in ["supervisor","witnesses"]: add_edit(name,value=existing[name] if existing else "")
        incident_type=QComboBox(); incident_type.addItems(INCIDENT_TYPES); form.addRow("Type of Incident:",incident_type)
        description=QTextEdit(); description.setMinimumHeight(100); description.setPlainText(safe(existing["description"]) if existing else ""); form.addRow("Incident Description:",description)
        immediate=QTextEdit(); immediate.setMinimumHeight(70); immediate.setPlainText(safe(existing["immediate_action"]) if existing else ""); form.addRow("Immediate Action:",immediate)
        # People involved
        people_box=QGroupBox("People Involved"); people_layout=QVBoxLayout(people_box); people_rows=[]
        def add_person(data=None):
            roww=QHBoxLayout(); n=QLineEdit(safe((data or {}).get("name",""))); role=QComboBox(); role.addItems(["Injured Person","Involved Person","Witness","Other"]); role.setCurrentText(safe((data or {}).get("role","Involved Person"))); d=QLineEdit(safe((data or {}).get("designation",""))); i=QLineEdit(safe((data or {}).get("id",""))); rem=QPushButton("Remove")
            n.setPlaceholderText("Person involved"); d.setPlaceholderText("Designation"); i.setPlaceholderText("ID"); roww.addWidget(n,2); roww.addWidget(role,2); roww.addWidget(d,2); roww.addWidget(i,1); roww.addWidget(rem); wrap=QWidget(); wrap.setLayout(roww); people_layout.addWidget(wrap); people_rows.append((wrap,n,role,d,i)); rem.clicked.connect(lambda: (people_rows.remove(next(x for x in people_rows if x[0] is wrap)), wrap.deleteLater()))
        addp=QPushButton("+ Add Person"); people_layout.addWidget(addp); addp.clicked.connect(lambda:add_person())
        if existing:
            try:
                for x in json.loads(safe(existing["people_details"]) or "[]"): add_person(x)
            except Exception: pass
        if not people_rows: add_person()
        form.addRow(people_box)
        # Equipment
        equipment_box=QGroupBox("Equipment"); equipment_layout=QVBoxLayout(equipment_box); equipment_rows=[]
        def add_equipment(data=None):
            roww=QHBoxLayout(); n=QLineEdit(safe((data or {}).get("name", ""))); i=QLineEdit(safe((data or {}).get("id", ""))); sp=QLineEdit(safe((data or {}).get("specification", ""))); rem=QPushButton("Remove")
            n.setPlaceholderText("Equipment"); i.setPlaceholderText("ID"); sp.setPlaceholderText("Specification"); roww.addWidget(n,2); roww.addWidget(i,1); roww.addWidget(sp,3); roww.addWidget(rem); wrap=QWidget(); wrap.setLayout(roww); equipment_layout.addWidget(wrap); equipment_rows.append((wrap,n,i,sp)); rem.clicked.connect(lambda: (equipment_rows.remove(next(x for x in equipment_rows if x[0] is wrap)), wrap.deleteLater()))
        adde=QPushButton("+ Add Equipment"); equipment_layout.addWidget(adde); adde.clicked.connect(lambda:add_equipment())
        if existing:
            try:
                for x in json.loads(safe(existing["equipment_details"]) or "[]"): add_equipment(x)
            except Exception: pass
        if not equipment_rows: add_equipment()
        form.addRow(equipment_box)
        # Repeatable incident detail sections
        ENV_CATEGORIES=["Spillage","Oil Spill","Chemical Spill","Fuel Spill","Diesel Spill","Hydrocarbon Spill","Water Discharge","Sewage Discharge","Air Emission","Dust Emission","Smoke / Fumes","Gas Release","Waste","Hazardous Waste","Non-Hazardous Waste","Soil Contamination","Water Contamination","Noise","Odour","Radiation","Resource Consumption","Energy Use","Water Use","Other"]
        def repeat_section(title, fields, old_key):
            box=QGroupBox(title); lay=QVBoxLayout(box); rows=[]
            def add(data=None):
                data=data or {}; h=QHBoxLayout(); widgets=[]
                for key,label in fields:
                    if key=="category":
                        w=QComboBox(); w.addItems(ENV_CATEGORIES); w.setCurrentText(safe(data.get(key,"")))
                    else:
                        w=QLineEdit(safe(data.get(key,""))); w.setPlaceholderText(label)
                    widgets.append(w); h.addWidget(w,3 if key in {"details","name"} else 2)
                rm=QPushButton("Remove"); h.addWidget(rm); wrap=QWidget(); wrap.setLayout(h); lay.addWidget(wrap); rows.append((wrap,widgets)); rm.clicked.connect(lambda: (rows.remove(next(x for x in rows if x[0] is wrap)),wrap.deleteLater()))
            plus=QPushButton("+ Add"); lay.addWidget(plus); plus.clicked.connect(lambda:add())
            vals=[]
            if existing:
                try: vals=json.loads(safe(existing["investigation_details"]) or "{}").get(old_key,[])
                except Exception: vals=[]
            for v in vals: add(v)
            if not rows: add()
            form.addRow(box); return rows
        environmental_rows=repeat_section("Environmental",[("category","Category (e.g. Spillage)"),("quantity","Quantity / Unit")],"environmental_items")
        property_rows=repeat_section("Property / Asset",[("name","Property / Asset Name"),("cost","Cost of Damage")],"property_items")
        procedure_rows=repeat_section("Procedure / Reference",[("reference","Procedure / Reference"),("details","Details")],"procedure_items")
        add_edit("organization","Organisation",existing["organization"] if existing else "")
        add_edit("report_reference","Report / Reference No.",existing["report_reference"] if existing else "")
        method=QComboBox(); method.addItems(INVESTIGATION_METHODS)
        if existing: method.setCurrentText(safe(existing["investigation_method"]))
        # Lock the selected investigation method so the method-specific content cannot
        # be changed accidentally.  Unlock is explicit and reversible.
        method_row=QHBoxLayout(); method_row.addWidget(method,1)
        method_lock=QPushButton("Lock Method"); method_lock.setMinimumHeight(34); method_lock.setMinimumWidth(125)
        method_state=QLabel("Unlocked")
        method_state.setStyleSheet("font-weight:bold;color:#8A6D1D;")
        method_row.addWidget(method_lock); method_row.addWidget(method_state)
        method_wrap=QWidget(); method_wrap.setLayout(method_row); form.addRow("Investigation Method:",method_wrap)
        method_locked=[False]
        def toggle_method_lock():
            method_locked[0]=not method_locked[0]
            method.setEnabled(not method_locked[0])
            if method_locked[0]:
                method_lock.setText("Unlock Method"); method_state.setText("Locked"); method_state.setStyleSheet("font-weight:bold;color:#B00020;")
            else:
                method_lock.setText("Lock Method"); method_state.setText("Unlocked"); method_state.setStyleSheet("font-weight:bold;color:#8A6D1D;")
        method_lock.clicked.connect(toggle_method_lock)
        procedure_box=QGroupBox("Investigation Procedure"); procedure_layout=QVBoxLayout(procedure_box); procedure_form=QFormLayout(); procedure_layout.addLayout(procedure_form); form.addRow(procedure_box); proc_widgets=[]
        def clear_proc():
            while procedure_form.count():
                item=procedure_form.takeAt(0); w=item.widget()
                if w:w.deleteLater()
            proc_widgets.clear()
        def add_proc(label,name,multi=True,value=""):
            w=QTextEdit() if multi else QLineEdit(); w.setPlainText(safe(value)) if multi else w.setText(safe(value));
            if multi:w.setMinimumHeight(65)
            proc_widgets.append((name,w)); procedure_form.addRow(label+":",w); return w
        def build_proc(_=None):
            clear_proc(); m=method.currentText(); old={}
            try: old=json.loads(safe(existing["investigation_details"]) or "{}") if existing else {}
            except Exception: old={}
            if m=="ICAM":
                for lab,key in [("Timeline / Event Sequence","timeline"),("Failed or Missing Defenses / Barriers","failed_defenses"),("Immediate Causes","immediate_causes"),("Contributing Conditions","contributing_conditions"),("Organizational Factors","organizational_factors"),("Root Cause","root_cause_analysis")]: add_proc(lab,key,value=old.get(key,""))
            elif m=="5 Why Analysis":
                for i in range(1,6): add_proc(f"Why {i}",f"why{i}",False,old.get(f"why{i}",""))
            elif m=="Fishbone / Ishikawa":
                # Structured Fishbone / Ishikawa RCA workspace. The category lists and
                # guidance below are based on the supplied Fishbone RCA structure.
                fish_box=QGroupBox("Fishbone / Ishikawa Analysis")
                fish_lay=QVBoxLayout(fish_box)
                fish_lay.addWidget(QLabel("Identify possible causes under each factor. Select a suggested cause or choose Other and add a specific cause."))
                fish_rows=[]
                fish_options={
                    "People": ["Lack of training","Expired / not evident training","Lack of experience","Fatigue","Human error","Poor communication","Inadequate supervision","Lack of competency","Other"],
                    "Machine / Equipment": ["Equipment failure","Poor maintenance","Incorrect settings","Defective component","Lack of guarding","Calibration problem","Other"],
                    "Method": ["Incorrect procedure","SOP not available","SOP not followed","Poor work planning","Inadequate risk assessment","Incorrect sequence of work","Other"],
                    "Material": ["Wrong material","Defective material","Poor quality","Incorrect specification","Contamination","Improper storage","Other"],
                    "Measurement": ["Incorrect measurement","Uncalibrated instrument","Wrong inspection method","Incorrect data","Inadequate monitoring","Other"],
                    "Environment": ["Temperature","Humidity","Lighting","Noise","Dust","Poor housekeeping","Congestion","Weather conditions","Other"],
                }
                saved_fish=old.get("fishbone",{}) if isinstance(old.get("fishbone",{}),dict) else {}
                for cat in fish_options:
                    for item in (saved_fish.get(cat,[]) or []):
                        pass
                fish_headers=QLabel("Factor | Suggested Cause | Specific Cause / Explanation | Evidence / Finding | Attachment")
                fish_headers.setStyleSheet("font-weight:bold;color:#17365D;")
                fish_lay.addWidget(fish_headers)
                fish_rows_box=QVBoxLayout()
                fish_lay.addLayout(fish_rows_box)
                # Keep every newly-added Fishbone cause inside this dedicated container.
                # This prevents rows from being inserted into the following Fishbone sections.
                def add_fish_row(category="People", cause="", detail="", evidence="", attachments=None):
                    attachments=list(attachments or [])
                    h=QHBoxLayout(); cat=QComboBox(); cat.addItems(list(fish_options.keys())+["Other"]); cat.setCurrentText(category if category in fish_options else "Other")
                    cause_cb=QComboBox(); cause_cb.setEditable(True); cause_cb.addItems(fish_options.get(cat.currentText(),["Other"])); cause_cb.setCurrentText(cause)
                    detail_e=QLineEdit(detail); detail_e.setPlaceholderText("Specific cause / explanation")
                    evidence_e=QLineEdit(evidence); evidence_e.setPlaceholderText("Evidence / finding")
                    attach_files=[]; attach_files.extend(attachments)
                    attach_lab=QLabel("; ".join(Path(x).name for x in attach_files) if attach_files else "No attachment"); attach_lab.setWordWrap(True)
                    attach=QPushButton("Add Attachment")
                    rm=QPushButton("Remove")
                    h.addWidget(cat,2); h.addWidget(cause_cb,3); h.addWidget(detail_e,3); h.addWidget(evidence_e,3); h.addWidget(attach,2); h.addWidget(rm,1)
                    wrap=QWidget(); wrap.setLayout(h); fish_rows_box.addWidget(wrap); fish_rows.append((wrap,cat,cause_cb,detail_e,evidence_e,attach_files,attach_lab))
                    # Put the filename label immediately below this cause row.
                    fish_rows_box.addWidget(attach_lab)
                    attach_lab.setStyleSheet("color:#555;font-size:11px;padding-left:8px;")
                    def pick_fish_attachment():
                        paths,_=QFileDialog.getOpenFileNames(dialog,"Select Fishbone Cause Evidence","","Evidence Files (*)")
                        if paths:
                            attach_files.clear(); attach_files.extend(paths); attach_lab.setText("; ".join(Path(x).name for x in attach_files))
                    attach.clicked.connect(pick_fish_attachment)
                    def refresh_causes(_=None):
                        current=cause_cb.currentText(); cause_cb.blockSignals(True); cause_cb.clear(); cause_cb.addItems(fish_options.get(cat.currentText(),["Other"])); cause_cb.setCurrentText(current); cause_cb.blockSignals(False)
                    cat.currentTextChanged.connect(refresh_causes)
                    rm.clicked.connect(lambda: remove_fish_row(wrap,attach_lab))
                def remove_fish_row(wrap,attach_lab=None):
                    item=next((x for x in fish_rows if x[0] is wrap),None)
                    if item:
                        fish_rows.remove(item); wrap.deleteLater()
                        if attach_lab: attach_lab.deleteLater()
                saved_rows=[]
                for cat,items in saved_fish.items():
                    for x in items or []:
                        if isinstance(x,dict): saved_rows.append((cat,x.get("cause",""),x.get("detail",""),x.get("evidence",""),x.get("attachments",[]) or []))
                if saved_rows:
                    for x in saved_rows: add_fish_row(*x)
                else:
                    add_fish_row()
                add_fish=QPushButton("+ Add Fishbone Cause")
                add_fish.setMinimumHeight(36)
                fish_lay.addWidget(add_fish)
                # QAbstractButton.clicked emits a boolean; use a no-argument lambda so
                # that the signal value cannot be mistaken for the cause category.
                add_fish.clicked.connect(lambda checked=False: add_fish_row())
                fish_lay.addWidget(QLabel("Fishbone factors: People, Machine / Equipment, Method, Material, Measurement and Environment. Use Add for additional applicable causes."))
                # Supporting Fishbone RCA sections from the supplied RCA flow. These are
                # only shown when Fishbone is selected, so other investigation methods keep
                # their existing UI unchanged.
                fish_problem=QTextEdit(); fish_problem.setMinimumHeight(70); fish_problem.setPlainText(safe(old.get("problem_statement",""))); fish_lay.addWidget(QLabel("Problem Statement (what went wrong; do not put the suspected root cause here)")); fish_lay.addWidget(fish_problem)
                team_box=QGroupBox("Investigation Team"); team_lay=QVBoxLayout(team_box); team_rows=[]
                def add_team(data=None):
                    data=data or {}; h=QHBoxLayout(); n=QLineEdit(safe(data.get("name",""))); pos=QLineEdit(safe(data.get("position",""))); dep=QLineEdit(safe(data.get("department",""))); role=QLineEdit(safe(data.get("role",""))); rm=QPushButton("Remove")
                    n.setPlaceholderText("Name"); pos.setPlaceholderText("Position"); dep.setPlaceholderText("Department"); role.setPlaceholderText("Role in Investigation")
                    h.addWidget(n,2); h.addWidget(pos,2); h.addWidget(dep,2); h.addWidget(role,2); h.addWidget(rm); w=QWidget(); w.setLayout(h); team_lay.addWidget(w); team_rows.append((w,n,pos,dep,role)); rm.clicked.connect(lambda:remove_team(w))
                def remove_team(w):
                    item=next((x for x in team_rows if x[0] is w),None)
                    if item: team_rows.remove(item); w.deleteLater()
                team_plus=QPushButton("+ Add Investigation Team Member"); team_lay.addWidget(team_plus); team_plus.clicked.connect(lambda:add_team())
                for x in (old.get("investigation_team",[]) or []): add_team(x)
                if not team_rows: add_team()
                fish_lay.addWidget(team_box)
                timeline_box=QGroupBox("Timeline of Events"); timeline_lay=QVBoxLayout(timeline_box); timeline_rows=[]
                def add_timeline(data=None):
                    data=data or {}; h=QHBoxLayout(); tm=QLineEdit(safe(data.get("time",""))); ev=QLineEdit(safe(data.get("event",""))); rm=QPushButton("Remove"); tm.setPlaceholderText("Time"); ev.setPlaceholderText("Event / What happened")
                    h.addWidget(tm,1); h.addWidget(ev,5); h.addWidget(rm); w=QWidget(); w.setLayout(h); timeline_lay.addWidget(w); timeline_rows.append((w,tm,ev)); rm.clicked.connect(lambda:remove_timeline(w))
                def remove_timeline(w):
                    item=next((x for x in timeline_rows if x[0] is w),None)
                    if item: timeline_rows.remove(item); w.deleteLater()
                tl_plus=QPushButton("+ Add Timeline Event"); timeline_lay.addWidget(tl_plus); tl_plus.clicked.connect(lambda:add_timeline())
                for x in (old.get("timeline_events",[]) or []): add_timeline(x)
                if not timeline_rows: add_timeline()
                fish_lay.addWidget(timeline_box)
                verify_box=QGroupBox("Root Cause Verification"); verify_lay=QVBoxLayout(verify_box); verify_rows=[]
                def add_verify(data=None):
                    data=data or {}; h=QHBoxLayout(); cause=QLineEdit(safe(data.get("cause",""))); evidence=QLineEdit(safe(data.get("evidence",""))); verified=QComboBox(); verified.addItems(["Yes","No","Pending"]); verified.setCurrentText(safe(data.get("verified","Pending")) or "Pending"); person=QLineEdit(safe(data.get("responsible",""))); rm=QPushButton("Remove")
                    cause.setPlaceholderText("Potential Cause"); evidence.setPlaceholderText("Evidence / Investigation Finding"); person.setPlaceholderText("Responsible for Verification")
                    h.addWidget(cause,2); h.addWidget(evidence,3); h.addWidget(verified,1); h.addWidget(person,2); h.addWidget(rm); w=QWidget(); w.setLayout(h); verify_lay.addWidget(w); verify_rows.append((w,cause,evidence,verified,person)); rm.clicked.connect(lambda:remove_verify(w))
                def remove_verify(w):
                    item=next((x for x in verify_rows if x[0] is w),None)
                    if item: verify_rows.remove(item); w.deleteLater()
                vr_plus=QPushButton("+ Add Root Cause Verification"); verify_lay.addWidget(vr_plus); vr_plus.clicked.connect(lambda:add_verify())
                for x in (old.get("root_cause_verification",[]) or []): add_verify(x)
                if not verify_rows: add_verify()
                fish_lay.addWidget(verify_box)
                fish_effective=QTextEdit(); fish_effective.setMinimumHeight(60); fish_effective.setPlainText(safe(old.get("effectiveness_verification",""))); fish_lay.addWidget(QLabel("Effectiveness Verification")); fish_lay.addWidget(fish_effective)
                fish_lessons=QTextEdit(); fish_lessons.setMinimumHeight(60); fish_lessons.setPlainText(safe(old.get("lessons_learned",""))); fish_lay.addWidget(QLabel("Lessons Learned")); fish_lay.addWidget(fish_lessons)
                fish_conclusion=QTextEdit(); fish_conclusion.setMinimumHeight(60); fish_conclusion.setPlainText(safe(old.get("conclusion",""))); fish_lay.addWidget(QLabel("Conclusion")); fish_lay.addWidget(fish_conclusion)
                proc_widgets.extend([
                    ("fish_problem_statement", fish_problem),
                    ("fish_investigation_team", team_rows),
                    ("fish_timeline_events", timeline_rows),
                    ("fish_root_cause_verification", verify_rows),
                    ("fish_effectiveness_verification", fish_effective),
                    ("fish_lessons_learned", fish_lessons),
                    ("fish_conclusion", fish_conclusion),
                ])
                # IMPORTANT: the Fishbone workspace belongs to the investigation
                # procedure container.  Putting it directly on the outer form caused
                # old Fishbone panels to remain when another method was selected.
                procedure_form.addRow(fish_box)
                # Preserve any legacy Fishbone text when an older record is edited.
                if not saved_rows:
                    legacy_map={"People":"people","Machine / Equipment":"machine","Method":"method","Material":"material","Environment":"environment","Measurement":"management"}
                    for cat,key in legacy_map.items():
                        legacy=old.get(key,"")
                        if legacy: add_fish_row(cat,"Other",legacy,"")
                proc_widgets.append(("fishbone_rows", fish_rows))
            elif m=="Barrier Analysis":
                for lab,key in [("Hazard / Threat","hazard"),("Top Event","top_event"),("Required Barriers","required_barriers"),("Failed / Missing Barriers","failed_barriers"),("Recovery / Mitigation","mitigation")]: add_proc(lab,key,value=old.get(key,""))
            elif m=="Bow-Tie Analysis":
                for lab,key in [("Threats","threats"),("Top Event","top_event"),("Preventive Barriers","preventive_barriers"),("Consequences","consequences"),("Mitigative Barriers","mitigative_barriers")]: add_proc(lab,key,value=old.get(key,""))
            elif m=="Fault Tree Analysis":
                for lab,key in [("Top Event","top_event"),("Intermediate Events","intermediate_events"),("Basic Events","basic_events"),("Logic / Gate Analysis","logic_analysis")]: add_proc(lab,key,value=old.get(key,""))
            elif m=="Causal Tree":
                for lab,key in [("Event","event"),("Causal Sequence","causal_sequence"),("Direct Causes","direct_causes_method"),("Contributing Causes","contributing_causes"),("Root Causes","root_causes_method")]: add_proc(lab,key,value=old.get(key,""))
            else: add_proc("Investigation Procedure / Notes","summary",value=old.get("summary",""))
        method.currentTextChanged.connect(build_proc); build_proc()
        # Multiple RCA fields
        def multi_cause_group(title,key):
            box=QGroupBox(title); lay=QVBoxLayout(box); rows=[]
            def add(value=""):
                h=QHBoxLayout(); e=QLineEdit(safe(value)); e.setPlaceholderText(title); rm=QPushButton("Remove"); h.addWidget(e,1); h.addWidget(rm); wrap=QWidget(); wrap.setLayout(h); lay.addWidget(wrap); rows.append((wrap,e)); rm.clicked.connect(lambda: (rows.remove(next(x for x in rows if x[0] is wrap)),wrap.deleteLater()))
            plus=QPushButton("+ Add"); lay.addWidget(plus); plus.clicked.connect(lambda:add());
            vals=[]
            if existing:
                try: vals=json.loads(safe(existing["investigation_details"]) or "{}").get(key,[])
                except Exception: vals=[]
            if isinstance(vals,str): vals=[vals]
            for v in vals: add(v)
            if not rows: add()
            form.addRow(box); return rows
        direct_rows=multi_cause_group("Direct Cause", "direct_causes")
        contrib_rows=multi_cause_group("Contributing Factors", "contributing_factors_list")
        root_rows=multi_cause_group("Root Cause", "root_causes_list")
        # Corrective action plan
        ca_box=QGroupBox("Corrective Action Plan"); ca_lay=QVBoxLayout(ca_box); ca_rows=[]
        def add_ca(data=None):
            data=data or {}; h=QHBoxLayout(); action=QLineEdit(safe(data.get("action",""))); resp=QLineEdit(safe(data.get("responsible",""))); target=QLineEdit(safe(data.get("target_date",""))); status=QComboBox(); status.addItems(["Open","Closed"]); status.setCurrentText(safe(data.get("status","Open")) or "Open"); attach=[]; lab=QLabel("No closeout evidence"); attach_btn=QPushButton("Attach"); rem=QPushButton("Remove")
            action.setPlaceholderText("Action"); resp.setPlaceholderText("Responsible Person"); target.setPlaceholderText("Target Date");
            h.addWidget(action,3); h.addWidget(resp,2); h.addWidget(target,1); h.addWidget(status,1); h.addWidget(attach_btn); h.addWidget(rem); wrap=QWidget(); wrap.setLayout(h); ca_lay.addWidget(wrap); h2=QHBoxLayout(); h2.addWidget(lab); wrap2=QWidget(); wrap2.setLayout(h2); ca_lay.addWidget(wrap2); ca_rows.append((wrap,action,resp,target,status,attach,lab))
            attach_btn.clicked.connect(lambda: choose_ca_files(attach,lab))
            rem.clicked.connect(lambda: remove_ca(wrap,wrap2))
        def choose_ca_files(store,lab):
            paths,_=QFileDialog.getOpenFileNames(dialog,"Select Closeout Evidence","","Evidence Files (*)")
            if paths: store.clear(); store.extend(paths); lab.setText("; ".join(Path(x).name for x in paths))
        def remove_ca(wrap,wrap2):
            item=next((x for x in ca_rows if x[0] is wrap),None)
            if item: ca_rows.remove(item); wrap.deleteLater(); wrap2.deleteLater()
        plus_ca=QPushButton("+ Add Corrective Action"); ca_lay.addWidget(plus_ca); plus_ca.clicked.connect(lambda:add_ca())
        old_actions=[]
        if existing:
            try: old_actions=json.loads(safe(existing["investigation_details"]) or "{}").get("corrective_actions",[])
            except Exception: old_actions=[]
        for a in old_actions: add_ca(a)
        if not ca_rows: add_ca()
        form.addRow(ca_box)
        # Witness statements: each statement can have its own attachment and remark.
        witness_box=QGroupBox("Witness Statements"); witness_lay=QVBoxLayout(witness_box); witness_rows=[]
        def add_witness(data=None):
            data=data or {}; n=QLineEdit(safe(data.get("name",""))); st=QTextEdit(); st.setFixedHeight(65); st.setPlainText(safe(data.get("statement",""))); rm=QLineEdit(safe(data.get("remark",""))); rm.setPlaceholderText("Remark"); files=[]; lab=QLabel(safe(data.get("attachment","")) or "No attachment"); attach=QPushButton("Attach"); remove=QPushButton("Remove")
            h=QHBoxLayout(); h.addWidget(n,2); h.addWidget(attach); h.addWidget(remove); wrap=QWidget(); wrap.setLayout(h); witness_lay.addWidget(wrap); witness_lay.addWidget(st); witness_lay.addWidget(rm); witness_lay.addWidget(lab); witness_rows.append((wrap,st,n,rm,files,lab))
            attach.clicked.connect(lambda: choose_witness(files,lab)); remove.clicked.connect(lambda: remove_witness(wrap,st,rm,lab))
        def choose_witness(store,lab):
            paths,_=QFileDialog.getOpenFileNames(dialog,"Select Witness Statement Attachment","","Files (*)")
            if paths: store.clear(); store.extend(paths); lab.setText("; ".join(Path(x).name for x in paths))
        def remove_witness(*widgets):
            for item in list(witness_rows):
                if item[0] is widgets[0]: witness_rows.remove(item); [x.deleteLater() for x in widgets]
        plus_w=QPushButton("+ Add Witness Statement"); witness_lay.addWidget(plus_w); plus_w.clicked.connect(lambda:add_witness())
        if existing:
            try:
                for x in json.loads(safe(existing["investigation_details"]) or "{}").get("witness_statements",[]): add_witness(x)
            except Exception: pass
        if not witness_rows: add_witness()
        form.addRow(witness_box)
        status=QComboBox(); status.addItems(["Draft","Open","Closed"]); status.setCurrentText(safe(existing["status"]) if existing else "Draft"); form.addRow("Investigation Status:",status)
        evidence_box=QGroupBox("Evidence Attachments"); evidence_lay=QVBoxLayout(evidence_box); evidence_rows=[]
        def add_evidence():
            h=QHBoxLayout(); file_store=[]; lab=QLabel("No file selected"); remark=QLineEdit(); remark.setPlaceholderText("Remark for this attachment"); choose=QPushButton("Add Attachment"); remove=QPushButton("Remove")
            h.addWidget(lab,2); h.addWidget(remark,3); h.addWidget(choose); h.addWidget(remove); wrap=QWidget(); wrap.setLayout(h); evidence_lay.addWidget(wrap); evidence_rows.append((wrap,file_store,remark,lab))
            def pick():
                paths,_=QFileDialog.getOpenFileNames(dialog,"Select Investigation Evidence","","Evidence Files (*)")
                if paths: file_store.clear(); file_store.extend(paths); lab.setText("; ".join(Path(x).name for x in paths))
            choose.clicked.connect(pick); remove.clicked.connect(lambda: remove_evidence(wrap))
        def remove_evidence(wrap):
            item=next((x for x in evidence_rows if x[0] is wrap),None)
            if item: evidence_rows.remove(item); wrap.deleteLater()
        plus_ev=QPushButton("+ Add Evidence Attachment"); evidence_lay.addWidget(plus_ev); plus_ev.clicked.connect(add_evidence); add_evidence(); form.addRow(evidence_box)
        # Keep action buttons permanently visible without needing to drag the form.
        action_bar=QHBoxLayout(); action_bar.addStretch()
        buttons=QDialogButtonBox(); save_draft=buttons.addButton("Save for Later",QDialogButtonBox.ButtonRole.AcceptRole); save_final=buttons.addButton("Save and Submit",QDialogButtonBox.ButtonRole.AcceptRole); save_exit=buttons.addButton("Save and Exit",QDialogButtonBox.ButtonRole.AcceptRole); cancel=buttons.addButton("Cancel",QDialogButtonBox.ButtonRole.RejectRole)
        for b in [save_draft,save_final,save_exit,cancel]: b.setMinimumHeight(38); b.setMinimumWidth(125)
        action_bar.addWidget(buttons); outer.insertLayout(1,action_bar); cancel.clicked.connect(dialog.reject)
        scroll.setWidget(content); outer.addWidget(scroll,1)
        def collect_rows(rows): return [e.text().strip() for _,e in rows if e.text().strip()]
        def save(mode):
            if not edits["location"].text().strip(): QMessageBox.warning(dialog,"Required","Location is required."); return
            try:
                number=safe(existing["number"]) if existing else next_number("HSE-INC","incidents")
                proc={}
                fishbone_rows_data=[]
                for name,w in proc_widgets:
                    if name=="fishbone_rows":
                        for _,cat,cause_cb,detail_e,evidence_e,attach_files,attach_lab in w:
                            if cat.currentText().strip() or cause_cb.currentText().strip() or detail_e.text().strip() or evidence_e.text().strip() or attach_files:
                                fishbone_rows_data.append({"category":cat.currentText().strip(),"cause":cause_cb.currentText().strip(),"detail":detail_e.text().strip(),"evidence":evidence_e.text().strip(),"attachments":list(attach_files)})
                    elif name=="fish_problem_statement":
                        proc["problem_statement"]=w.toPlainText()
                    elif name=="fish_investigation_team":
                        proc["investigation_team"]=[{"name":n.text().strip(),"position":pos.text().strip(),"department":dep.text().strip(),"role":role.text().strip()} for _,n,pos,dep,role in w if any(z.text().strip() for z in (n,pos,dep,role))]
                    elif name=="fish_timeline_events":
                        proc["timeline_events"]=[{"time":tm.text().strip(),"event":ev.text().strip()} for _,tm,ev in w if tm.text().strip() or ev.text().strip()]
                    elif name=="fish_root_cause_verification":
                        proc["root_cause_verification"]=[{"cause":cause.text().strip(),"evidence":evidence.text().strip(),"verified":verified.currentText(),"responsible":person.text().strip()} for _,cause,evidence,verified,person in w if any(z.text().strip() for z in (cause,evidence,person))]
                    elif name=="fish_effectiveness_verification":
                        proc["effectiveness_verification"]=w.toPlainText()
                    elif name=="fish_lessons_learned":
                        proc["lessons_learned"]=w.toPlainText()
                    elif name=="fish_conclusion":
                        proc["conclusion"]=w.toPlainText()
                    else:
                        proc[name]=(w.toPlainText() if isinstance(w,QTextEdit) else w.text())
                direct_vals=collect_rows(direct_rows); contrib_vals=collect_rows(contrib_rows); root_vals=collect_rows(root_rows)
                actions=[]
                for _,a,r,t,st,att,lab in ca_rows:
                    if a.text().strip() or r.text().strip() or t.text().strip() or att:
                        actions.append({"action":a.text().strip(),"responsible":r.text().strip(),"target_date":t.text().strip(),"status":st.currentText(),"attachments":[]})
                details=proc; details.update({"direct_causes":direct_vals,"contributing_factors_list":contrib_vals,"root_causes_list":root_vals,"corrective_actions":actions})
                if method.currentText()=="Fishbone / Ishikawa":
                    details["fishbone"]=fishbone_rows_data
                details["environmental_items"]=[{"category":widgets[0].currentText().strip(),"quantity":widgets[1].text().strip()} for _,widgets in environmental_rows if widgets[0].currentText().strip() or widgets[1].text().strip()]
                details["property_items"]=[{"name":widgets[0].text().strip(),"cost":widgets[1].text().strip()} for _,widgets in property_rows if any(w.text().strip() for w in widgets)]
                details["procedure_items"]=[{"reference":widgets[0].text().strip(),"details":widgets[1].text().strip()} for _,widgets in procedure_rows if any(w.text().strip() for w in widgets)]
                details["witness_statements"]=[]
                for _,st,n,rm,files,lab in witness_rows:
                    if n.text().strip() or st.toPlainText().strip() or files:
                        details["witness_statements"].append({"name":n.text().strip(),"statement":st.toPlainText().strip(),"remark":rm.text().strip(),"attachment":""})
                people=[{"name":n.text().strip(),"role":role.currentText(),"designation":d.text().strip(),"id":i.text().strip()} for _,n,role,d,i in people_rows if n.text().strip()]
                equipment=[{"name":n.text().strip(),"id":i.text().strip(),"specification":sp.text().strip()} for _,n,i,sp in equipment_rows if n.text().strip()]
                primary_corrective="\n".join(f"{i}. {x['action']}" for i,x in enumerate(actions,1))
                params=(today(),edits["incident_time"].text(),edits["location"].text(),edits["project"].text(),edits["company"].text(),edits["department"].text(),edits["activity"].text(),incident_type.currentText(),people[0]["name"] if people else "",people[0]["id"] if people else "",people[0]["designation"] if people else "",edits.get("supervisor",QLineEdit()).text() if "supervisor" in edits else "",edits.get("witnesses",QLineEdit()).text() if "witnesses" in edits else "",description.toPlainText(),immediate.toPlainText(),"", "", equipment[0]["name"] if equipment else "",method.currentText(),*(proc.get(f"why{i}","") for i in range(1,6)),"\n".join(direct_vals),"\n".join(contrib_vals),"\n".join(root_vals),primary_corrective,"",json.dumps(details,ensure_ascii=False),json.dumps(people,ensure_ascii=False),json.dumps(equipment,ensure_ascii=False),"", "", "", "", edits["organization"].text(), "", edits["report_reference"].text(),mode,datetime.now().isoformat())
                if existing:
                    db.execute("""UPDATE incidents SET incident_date=?,incident_time=?,location=?,project=?,company=?,department=?,activity=?,incident_type=?,person_involved=?,employee_id=?,designation=?,supervisor=?,witnesses=?,description=?,immediate_action=?,consequences=?,potential_consequences=?,equipment=?,investigation_method=?,why1=?,why2=?,why3=?,why4=?,why5=?,direct_cause=?,contributing_factors=?,root_cause=?,corrective_action=?,preventive_action=?,investigation_details=?,people_details=?,equipment_details=?,environmental_category=?,environmental_quantity=?,property_name=?,property_damage_cost=?,organization=?,procedure_reference=?,report_reference=?,status=?,created_at=? WHERE id=?""",params+(incident_id,))
                else:
                    db.execute("""INSERT INTO incidents (number,incident_date,incident_time,location,project,company,department,activity,incident_type,person_involved,employee_id,designation,supervisor,witnesses,description,immediate_action,consequences,potential_consequences,equipment,investigation_method,why1,why2,why3,why4,why5,direct_cause,contributing_factors,root_cause,corrective_action,preventive_action,investigation_details,people_details,equipment_details,environmental_category,environmental_quantity,property_name,property_damage_cost,organization,procedure_reference,report_reference,status,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(number,)+params)
                rid=incident_id or db.fetchone("SELECT id FROM incidents WHERE number=?",(number))["id"]
                db.execute("UPDATE incidents SET incident_title=? WHERE id=?", (edits["incident_title"].text().strip(), rid))
                # Save Fishbone cause attachments and replace temporary source paths with
                # the saved filenames so reports can display the actual evidence.
                if method.currentText()=="Fishbone / Ishikawa":
                    for frow in fishbone_rows_data:
                        srcs=frow.get("attachments",[]) or []
                        saved_names=[]
                        new_sources=[]
                        for src in srcs:
                            sp=Path(src)
                            if sp.exists():
                                new_sources.append(str(sp))
                            else:
                                # Existing Fishbone attachments are already stored in the
                                # incident attachment table; retain their filenames.
                                saved_names.append(sp.name)
                        if new_sources:
                            saved=copy_attachments(new_sources,number,"incident_attachments",rid,"Fishbone Cause")
                            saved_names.extend(Path(x).name for x in saved)
                        frow["attachments"]=saved_names
                    details["fishbone"]=fishbone_rows_data
                    db.execute("UPDATE incidents SET investigation_details=? WHERE id=?", (json.dumps(details, ensure_ascii=False), rid))
                for _,file_store,remark,lab in evidence_rows:
                    if file_store:
                        saved=copy_attachments(file_store,number,"incident_attachments",rid,"Evidence")
                        for saved_path in saved:
                            db.execute("UPDATE incident_attachments SET remark=? WHERE incident_id=? AND file_path=?",(remark.text().strip(),rid,saved_path))
                for idx, item in enumerate(witness_rows):
                    if idx >= len(details["witness_statements"]): continue
                    files=item[4]
                    if files:
                        saved=copy_attachments(files,number,"incident_attachments",rid,"Witness Statement")
                        details["witness_statements"][idx]["attachment"]="; ".join(Path(x).name for x in saved)
                        for saved_path in saved:
                            db.execute("UPDATE incident_attachments SET remark=? WHERE incident_id=? AND file_path=?",(item[3].text().strip(),rid,saved_path))
                # Save closeout evidence against the corrective-action plan items.
                for idx, item in enumerate(ca_rows):
                    if idx >= len(actions):
                        continue
                    files = item[5]
                    if files:
                        saved = copy_attachments(files, number, "incident_attachments", rid, "Closeout Evidence")
                        actions[idx]["attachments"] = [Path(x).name for x in saved]
                # Persist the final corrective-action attachment names without altering legacy fields.
                details["corrective_actions"] = actions
                db.execute("UPDATE incidents SET investigation_details=? WHERE id=?", (json.dumps(details, ensure_ascii=False), rid))
                dialog.accept(); refresh()
            except Exception as e: logging.exception("Incident save failed"); QMessageBox.critical(dialog,"Save Error",f"Unable to save investigation.\n\n{e}")
        save_draft.clicked.connect(lambda:save("Draft")); save_final.clicked.connect(lambda:save("Closed")); save_exit.clicked.connect(lambda:save("Draft")); dialog.exec()

    def audits(self):
        w,layout=self.page("Audit Register")
        toolbar=QHBoxLayout(); btns={}
        for text in ["+ New Audit","Delete Selected","Export CSV","Export Excel","Export PDF","Export Word"]:
            b=QPushButton(text); toolbar.addWidget(b); btns[text]=b
        layout.addLayout(toolbar)
        table=QTableWidget(); headers=["ID","Number","Date","Standard","Audit Type","Project","Auditor","Status","Attachments"]; table.setColumnCount(len(headers)); table.setHorizontalHeaderLabels(headers); table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents); table.horizontalHeader().setStretchLastSection(True); table.setSelectionBehavior(QAbstractItemView.SelectRows); layout.addWidget(table)
        def load():
            rows=db.fetchall("SELECT * FROM audits ORDER BY id DESC"); table.setRowCount(len(rows))
            for r,row in enumerate(rows):
                count=db.fetchone("SELECT COUNT(*) c FROM audit_attachments WHERE audit_id=?",(row["id"],))["c"]; vals=[row["id"],row["number"],row["audit_date"],row["standard"],row["audit_type"],row["project"],row["lead_auditor"],row["status"],count]
                for c,v in enumerate(vals):table.setItem(r,c,QTableWidgetItem(safe(v)))
        btns["+ New Audit"].clicked.connect(lambda:self.audit_form(load)); btns["Delete Selected"].clicked.connect(lambda:self.delete_selected_record(table,"audits","audit_attachments","Audit",load)); btns["Export CSV"].clicked.connect(lambda:self.export_csv("audits")); btns["Export Excel"].clicked.connect(lambda:self.export_excel("audits")); btns["Export PDF"].clicked.connect(lambda:self.export_pdf("audits")); btns["Export Word"].clicked.connect(lambda:self.export_docx("audits")); load()


    def audit_form(self, refresh):
        dialog = QDialog(self)
        dialog.setWindowTitle("New Audit")
        dialog.resize(1000, 900)
        dialog.setMinimumSize(920, 760)
        outer = QVBoxLayout(dialog)

        title = QLabel("AUDIT REGISTER - NEW AUDIT")
        title.setStyleSheet("font-size:20px;font-weight:bold;color:#17365D;padding:6px;")
        outer.addWidget(title)

        scroll_area = QScrollArea(); scroll_area.setWidgetResizable(True)
        content = QWidget(); form = QFormLayout(content)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        edits = {}

        def edit(name):
            e = QLineEdit(); edits[name] = e
            form.addRow(name.replace("_", " ").title() + ":", e)
            return e

        standard = QComboBox(); standard.addItems(["ISO 45001", "ISO 14001"])
        form.addRow("Standard:", standard)
        audit_type = QComboBox(); audit_type.addItems(AUDIT_TYPES)
        form.addRow("Audit Type:", audit_type)
        for name in ["project", "location", "department", "auditor", "lead_auditor", "auditee", "start_time", "end_time"]:
            edit(name)

        scope = QTextEdit(); scope.setMinimumHeight(70); form.addRow("Scope:", scope)
        objective = QTextEdit(); objective.setMinimumHeight(70); form.addRow("Objective:", objective)
        criteria = QTextEdit(); criteria.setMinimumHeight(70); form.addRow("Audit Criteria:", criteria)

        clause = QComboBox()
        sub_clause = QComboBox()
        form.addRow("ISO Clause:", clause)
        form.addRow("ISO Sub-Clause:", sub_clause)

        def load_clauses():
            clause.blockSignals(True)
            clause.clear()
            clause.addItems(list(ISO_CLAUSES.get(standard.currentText(), {}).keys()))
            clause.blockSignals(False)
            load_subclauses()

        def load_subclauses():
            sub_clause.clear()
            items = ISO_CLAUSES.get(standard.currentText(), {}).get(clause.currentText(), [])
            sub_clause.addItems(items)

        standard.currentTextChanged.connect(load_clauses)
        clause.currentTextChanged.connect(load_subclauses)
        load_clauses()

        requirement = QTextEdit(); requirement.setMinimumHeight(70); form.addRow("Requirement / Criterion:", requirement)
        finding_type = QComboBox(); finding_type.addItems(FINDING_TYPES); form.addRow("Finding Type:", finding_type)
        finding = QTextEdit(); finding.setMinimumHeight(90); form.addRow("Finding / Evidence:", finding)
        risk = QTextEdit(); risk.setMinimumHeight(70); form.addRow("Risk / Impact:", risk)
        corrective = QTextEdit(); corrective.setMinimumHeight(70); form.addRow("Corrective Action:", corrective)
        responsible = QLineEdit(); form.addRow("Responsible Person:", responsible)
        target = QLineEdit(); form.addRow("Target Date:", target)

        attachment_paths = []; attachment_label = QLabel("No evidence files selected."); attachment_label.setWordWrap(True)
        attach = QPushButton("Attach Audit Evidence Files"); attach.setMinimumHeight(36)
        def choose_files():
            paths, _ = QFileDialog.getOpenFileNames(
                dialog, "Select Audit Evidence", "",
                "Evidence Files (*.png *.jpg *.jpeg *.bmp *.pdf *.doc *.docx *.xls *.xlsx *.txt *.mp4 *.avi *.mov);;All Files (*)"
            )
            if paths:
                attachment_paths.clear(); attachment_paths.extend(paths)
                attachment_label.setText("\n".join(Path(p).name for p in paths))
        attach.clicked.connect(choose_files)
        form.addRow("Evidence / Attachments:", attach); form.addRow("Selected Files:", attachment_label)

        scroll_area.setWidget(content); outer.addWidget(scroll_area, 1)
        buttons = QDialogButtonBox()
        save_button = buttons.addButton("Save Audit", QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
        save_button.setMinimumHeight(40); cancel_button.setMinimumHeight(40)
        outer.addWidget(buttons); cancel_button.clicked.connect(dialog.reject)

        def save():
            try:
                number = next_number("HSE-AUD", "audits")
                db.execute("""
                    INSERT INTO audits (number,audit_date,audit_type,standard,project,location,department,
                        auditor,lead_auditor,auditee,scope,objective,criteria,start_time,end_time,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    number, today(), audit_type.currentText(), standard.currentText(), edits["project"].text(),
                    edits["location"].text(), edits["department"].text(), edits["auditor"].text(),
                    edits["lead_auditor"].text(), edits["auditee"].text(), scope.toPlainText(),
                    objective.toPlainText(), criteria.toPlainText(), edits["start_time"].text(),
                    edits["end_time"].text(), datetime.now().isoformat()
                ))
                audit = db.fetchone("SELECT id FROM audits WHERE number=?", (number,))
                db.execute("""
                    INSERT INTO audit_findings (audit_id,clause,sub_clause,requirement,finding_type,
                        observation,evidence,risk_impact,corrective_action,responsible,target_date)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    audit["id"], clause.currentText(), sub_clause.currentText(), requirement.toPlainText(),
                    finding_type.currentText(), finding.toPlainText(), finding.toPlainText(), risk.toPlainText(),
                    corrective.toPlainText(), responsible.text(), target.text()
                ))
                if attachment_paths:
                    copy_attachments(attachment_paths, number, "audit_attachments", audit["id"])
                dialog.accept(); refresh()
            except Exception as e:
                logging.exception("Audit save failed")
                QMessageBox.critical(dialog, "Save Error", f"Unable to save audit.\n\n{e}")

        save_button.clicked.connect(save)
        dialog.exec()

    def capa(self):
        w,layout=self.page("CAPA Register")
        toolbar=QHBoxLayout(); btns={}
        for text in ["+ New CAPA","Delete Selected","Export CSV","Export Excel","Export PDF","Export Word"]:
            b=QPushButton(text); toolbar.addWidget(b); btns[text]=b
        layout.addLayout(toolbar)
        table=QTableWidget(); headers=["ID","Number","Source","Reference","Priority","Responsible","Target Date","Status","Overdue","Attachments"]; table.setColumnCount(len(headers)); table.setHorizontalHeaderLabels(headers); table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents); table.horizontalHeader().setStretchLastSection(True); table.setSelectionBehavior(QAbstractItemView.SelectRows); layout.addWidget(table)
        def load():
            rows=db.fetchall("SELECT * FROM capa ORDER BY id DESC"); table.setRowCount(len(rows))
            for r,row in enumerate(rows):
                count=db.fetchone("SELECT COUNT(*) c FROM capa_attachments WHERE capa_id=?",(row["id"],))["c"]; vals=[row["id"],row["number"],row["source"],row["reference_number"],row["priority"],row["responsible"],row["target_date"],row["status"],"OVERDUE" if overdue(row["target_date"],row["status"]) else "",count]
                for c,v in enumerate(vals):table.setItem(r,c,QTableWidgetItem(safe(v)))
        btns["+ New CAPA"].clicked.connect(lambda:self.capa_form(load)); btns["Delete Selected"].clicked.connect(lambda:self.delete_selected_record(table,"capa","capa_attachments","CAPA",load)); btns["Export CSV"].clicked.connect(lambda:self.export_csv("capa")); btns["Export Excel"].clicked.connect(lambda:self.export_excel("capa")); btns["Export PDF"].clicked.connect(lambda:self.export_pdf("capa")); btns["Export Word"].clicked.connect(lambda:self.export_docx("capa")); load()


    def capa_form(self, refresh):
        dialog = QDialog(self)
        dialog.setWindowTitle("New CAPA")
        dialog.resize(950, 850)
        dialog.setMinimumSize(900, 720)
        outer = QVBoxLayout(dialog)

        title = QLabel("CORRECTIVE AND PREVENTIVE ACTION (CAPA)")
        title.setStyleSheet("font-size:20px;font-weight:bold;color:#17365D;padding:6px;")
        outer.addWidget(title)

        scroll_area = QScrollArea(); scroll_area.setWidgetResizable(True)
        content = QWidget(); form = QFormLayout(content)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)

        source = QComboBox(); source.addItems(CAPA_SOURCES); form.addRow("Source:", source)
        reference = QLineEdit(); form.addRow("Reference Number:", reference)
        finding = QTextEdit(); finding.setMinimumHeight(90); form.addRow("Finding:", finding)
        root = QTextEdit(); root.setMinimumHeight(90); form.addRow("Root Cause:", root)
        corrective = QTextEdit(); corrective.setMinimumHeight(90); form.addRow("Corrective Action:", corrective)
        preventive = QTextEdit(); preventive.setMinimumHeight(90); form.addRow("Preventive Action:", preventive)
        responsible = QLineEdit(); form.addRow("Responsible Person:", responsible)
        priority = QComboBox(); priority.addItems(PRIORITIES); form.addRow("Priority:", priority)
        target = QLineEdit(); form.addRow("Target Date:", target)
        status = QComboBox(); status.addItems(STATUSES); form.addRow("Status:", status)
        verification = QTextEdit(); verification.setMinimumHeight(70); form.addRow("Verification:", verification)
        evidence = QTextEdit(); evidence.setMinimumHeight(70); form.addRow("Closeout Evidence:", evidence)

        attachment_paths = []; attachment_label = QLabel("No evidence files selected."); attachment_label.setWordWrap(True)
        attach = QPushButton("Attach CAPA Evidence Files"); attach.setMinimumHeight(36)
        def choose_files():
            paths, _ = QFileDialog.getOpenFileNames(
                dialog, "Select CAPA Evidence", "",
                "Evidence Files (*.png *.jpg *.jpeg *.bmp *.pdf *.doc *.docx *.xls *.xlsx *.txt *.mp4 *.avi *.mov);;All Files (*)"
            )
            if paths:
                attachment_paths.clear(); attachment_paths.extend(paths)
                attachment_label.setText("\n".join(Path(p).name for p in paths))
        attach.clicked.connect(choose_files)
        form.addRow("Evidence / Attachments:", attach); form.addRow("Selected Files:", attachment_label)

        scroll_area.setWidget(content); outer.addWidget(scroll_area, 1)
        buttons = QDialogButtonBox()
        save_button = buttons.addButton("Save CAPA", QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
        save_button.setMinimumHeight(40); cancel_button.setMinimumHeight(40)
        outer.addWidget(buttons); cancel_button.clicked.connect(dialog.reject)

        def save():
            try:
                number = next_number("HSE-CAPA", "capa")
                db.execute("""
                    INSERT INTO capa (number,source,reference_number,finding,root_cause,corrective_action,
                        preventive_action,responsible,priority,target_date,status,verification,closeout_evidence,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    number, source.currentText(), reference.text(), finding.toPlainText(), root.toPlainText(),
                    corrective.toPlainText(), preventive.toPlainText(), responsible.text(), priority.currentText(),
                    target.text(), status.currentText(), verification.toPlainText(), evidence.toPlainText(),
                    datetime.now().isoformat()
                ))
                row = db.fetchone("SELECT id FROM capa WHERE number=?", (number,))
                if row and attachment_paths:
                    copy_attachments(attachment_paths, number, "capa_attachments", row["id"])
                dialog.accept(); refresh()
            except Exception as e:
                logging.exception("CAPA save failed")
                QMessageBox.critical(dialog, "Save Error", f"Unable to save CAPA.\n\n{e}")

        save_button.clicked.connect(save)
        dialog.exec()

    def reports(self):
        w, layout = self.page("Reports & Export")
        banner = QHBoxLayout()
        logo_path = db.setting("company_logo", "")
        logo = QLabel()
        logo.setFixedSize(90, 70)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        if logo_path and Path(logo_path).exists():
            pix = QPixmap(logo_path)
            logo.setPixmap(pix.scaled(80, 65, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        banner.addWidget(logo)
        company_label = QLabel(f"<b>{safe(company_name())}</b><br>Reports & Export<br>Project: {safe(db.setting('project_name',''))}")
        company_label.setStyleSheet("font-size:16px;padding:6px;")
        banner.addWidget(company_label, 1)
        open_folder = QPushButton("Open Reports Folder")
        open_folder.clicked.connect(self.open_reports_folder)
        banner.addWidget(open_folder)
        layout.addLayout(banner)

        info = QLabel("Select a register and format. The report is extracted from the saved HSE database and can be saved to any folder on the computer.")
        info.setWordWrap(True)
        layout.addWidget(info)

        for title, table_name in [
            ("HSE Inspection / Observation Report", "observations"),
            ("Incident Investigation Register Report", "incidents"),
            ("Audit Register Report", "audits"),
            ("CAPA Register Report", "capa")
        ]:
            box = QGroupBox(title)
            row = QHBoxLayout(box)
            for label, handler in [
                ("Download CSV", lambda t=table_name: self.export_csv(t)),
                ("Download Excel", lambda t=table_name: self.export_excel(t)),
                ("Download PDF", lambda t=table_name: self.export_pdf(t)),
                ("Download Word", lambda t=table_name: self.export_docx(t))
            ]:
                b = QPushButton(label)
                b.setMinimumHeight(38)
                row.addWidget(b)
                b.clicked.connect(handler)
            layout.addWidget(box)

        incident_box = QGroupBox("Professional Incident Investigation Word Report")
        ir = QHBoxLayout(incident_box)
        incident_combo = QComboBox()
        incidents = db.fetchall("SELECT id,number FROM incidents ORDER BY id DESC")
        for row in incidents:
            incident_combo.addItem(safe(row["number"]), row["id"])
        ir.addWidget(QLabel("Investigation:"))
        ir.addWidget(incident_combo, 1)
        b = QPushButton("Download Investigation Report")
        b.setMinimumHeight(38)
        ir.addWidget(b)

        def generate():
            if incident_combo.count() == 0:
                QMessageBox.information(self, "No Incidents", "No incident investigations are available.")
                return
            incident_id = incident_combo.currentData()
            number = incident_combo.currentText()
            default_path = str(REPORT_DIR / f"{number}_Investigation_Report.docx")
            path, _ = QFileDialog.getSaveFileName(
                self, "Save Investigation Report", default_path, "Word Document (*.docx)"
            )
            if not path:
                return
            try:
                generate_incident_docx(int(incident_id), path)
                self.show_export_success(path)
            except Exception as e:
                logging.exception("Incident report failed")
                QMessageBox.critical(self, "Report Error", str(e))

        b.clicked.connect(generate)
        layout.addWidget(incident_box)
        layout.addStretch(1)

    def show_export_success(self, path):
        msg = QMessageBox(self)
        msg.setWindowTitle("Export Complete")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText("Report created successfully.")
        msg.setInformativeText(str(path))
        open_button = msg.addButton("Open Folder", QMessageBox.ButtonRole.ActionRole)
        msg.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        msg.exec()
        if msg.clickedButton() == open_button:
            self.open_path(Path(path).parent)

    def open_path(self, path):
        try:
            path = Path(path)
            if sys.platform.startswith("win"):
                os.startfile(str(path))
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception as e:
            logging.exception("Unable to open path")
            QMessageBox.warning(self, "Open Folder", f"Unable to open folder.\n\n{e}")

    def open_reports_folder(self):
        REPORT_DIR.mkdir(parents=True, exist_ok=True)
        self.open_path(REPORT_DIR)

    def _export_path(self, filename, title, file_filter):
        REPORT_DIR.mkdir(parents=True, exist_ok=True)
        return QFileDialog.getSaveFileName(self, title, str(REPORT_DIR / filename), file_filter)[0]

    @staticmethod
    def table_data_static(table_name, record_ids=None):
        allowed = {"observations", "incidents", "audits", "capa"}
        if table_name not in allowed:
            raise ValueError("Invalid report table.")
        if table_name == "observations":
            columns, data, _ = observation_export_rows(record_ids)
            return columns, data
        params = []
        where = ""
        if record_ids:
            placeholders = ",".join("?" for _ in record_ids)
            where = f" WHERE id IN ({placeholders})"
            params = list(record_ids)
        rows = db.fetchall(f"SELECT * FROM {table_name}{where} ORDER BY id DESC", params)
        if not rows:
            return [], []
        columns = list(rows[0].keys())
        attach_table = attachment_export_map(table_name)
        if attach_table and "Attachments" not in columns:
            columns.append("Attachments")
        data=[]
        id_key="id"
        for row in rows:
            values=[xml_safe(row[c]) for c in columns if c != "Attachments"]
            if attach_table:
                values.append(xml_safe(attachment_names(attach_table, row[id_key])))
            data.append(values)
        return columns, data

    def table_data(self, table_name, record_ids=None):
        return self.table_data_static(table_name, record_ids=record_ids)

    def selected_id(self, table, title):
        r=table.currentRow()
        if r < 0 or not table.item(r,0):
            QMessageBox.warning(self, title, "Select a record first.")
            return None
        return int(table.item(r,0).text())

    def delete_selected_record(self, table, table_name, attachment_table, title, refresh):
        record_id=self.selected_id(table, title)
        if record_id is None:
            return
        record=db.fetchone(f"SELECT number FROM {table_name} WHERE id=?", (record_id,))
        if not record:
            return
        if QMessageBox.question(self, "Confirm Delete", f"Delete {title.lower()} {record['number']} and its evidence?\n\nThis cannot be undone.") != QMessageBox.Yes:
            return
        for a in attachment_rows(attachment_table, record_id):
            try: Path(safe(a["file_path"])).unlink(missing_ok=True)
            except Exception: logging.exception("Unable to remove attachment")
        db.execute(f"DELETE FROM {table_name} WHERE id=?", (record_id,))
        refresh()

    def export_csv(self, table_name, record_ids=None):
        try:
            columns, data = self.table_data(table_name, record_ids)
            if not columns:
                QMessageBox.information(self, "No Data", "There is no data to export.")
                return
            path = self._export_path(f"{table_name}_report.csv", "Download CSV Report", "CSV Files (*.csv)")
            if not path: return
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer=csv.writer(f); writer.writerow(columns); writer.writerows(data)
            self.show_export_success(path)
        except Exception as e:
            logging.exception("CSV export failed"); QMessageBox.critical(self,"Export Error",str(e))

    def export_excel(self, table_name, record_ids=None):
        try:
            columns, data = self.table_data(table_name, record_ids)
            if not columns:
                QMessageBox.information(self, "No Data", "There is no data to export."); return
            path=self._export_path(f"{table_name}_report.xlsx","Download Excel Report","Excel Files (*.xlsx)")
            if not path: return
            wb=Workbook(); ws=wb.active; ws.title=table_name[:31]
            if table_name == "observations":
                preferred = ["S.No.", "Observation No.", "Observation Date", "Location", "Observation Type", "Category", "Observation Detail", "Action Taken", "Responsible", "Evidence", "Status"]
                # Professional three-row report header.  The layout is fixed to A1:K3:
                # A1:C3 = logo, D1:H3 = report heading, I1:K1 = report number,
                # I2:K2 = document number, I3:K3 = project name.
                ws.merge_cells("A1:C3")
                ws.merge_cells("D1:H3")
                ws.merge_cells("I1:K1")
                ws.merge_cells("I2:K2")
                ws.merge_cells("I3:K3")
                ws["D1"] = f"HSE OBSERVATION REPORT\n{company_name()}"
                report_no = f"{document_prefix()}-OBS-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                ws["I1"] = f"Report No.: {report_no}"
                ws["I2"] = f"Document No.: {document_prefix()}-OBS"
                ws["I3"] = f"Project: {db.setting('project_name', '')}"
                ws.append(preferred)
                obs_rows = db.fetchall("SELECT * FROM observations ORDER BY id DESC")
                if record_ids:
                    obs_rows = [r for r in obs_rows if r["id"] in record_ids]
                for serial, r in enumerate(obs_rows, 1):
                    vals=[serial, r["number"], r["obs_date"], r["location"], r["obs_type"], r["category"], r["observation"], r["corrective_action"], r["responsible"], "", r["status"]]
                    ws.append([xml_safe(v) for v in vals])
                    excel_row=ws.max_row
                    evcell=ws.cell(excel_row,10)
                    evcell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
                    image_files=[]; other_files=[]
                    for a in attachment_rows("observation_attachments", r["id"]):
                        fp=Path(safe(a["file_path"]))
                        if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp"}:
                            image_files.append(fp)
                        elif fp.exists():
                            other_files.append(fp.name)
                    # Show the first real image directly in the Evidence cell; all other evidence names remain visible.
                    if image_files:
                        try:
                            img=XLImage(str(image_files[0]))
                            max_w=145; max_h=62
                            ratio=min(max_w/max(img.width,1), max_h/max(img.height,1), 1)
                            img.width=max(1,int(img.width*ratio)); img.height=max(1,int(img.height*ratio))
                            img.anchor=f"J{excel_row}"
                            ws.add_image(img)
                            extra=[p.name for p in image_files[1:]] + other_files
                            evcell.value="; ".join(extra) if extra else "Photo evidence"
                        except Exception:
                            logging.exception("Unable to embed observation image in Excel")
                            evcell.value="; ".join([p.name for p in image_files] + other_files)
                    elif other_files:
                        evcell.value="; ".join(other_files)
                    else:
                        evcell.value="None"
                    ws.row_dimensions[excel_row].height=70
                # Professional register heading. The logo is fixed inside the merged A1:C3 area.
                logo_path=db.setting("company_logo", "")
                if logo_path and Path(logo_path).exists():
                    try:
                        logo=XLImage(logo_path)
                        logo.width=125; logo.height=62
                        logo.anchor="A1"
                        ws.add_image(logo)
                    except Exception:
                        logging.exception("Unable to embed company logo in Excel heading")
                # Header area is deliberately fixed to three rows; the register header begins on row 4.
                for r in range(1,4):
                    ws.row_dimensions[r].height = 24
                ws.row_dimensions[4].height = 30
                header_fill=PatternFill("solid", fgColor="D9E2F3")
                for cell_ref in ["A1","D1","I1","I2","I3"]:
                    c=ws[cell_ref]
                    c.font=copy(c.font); c.font=c.font.copy(bold=True, size=13 if cell_ref=="D1" else 10)
                    c.fill=header_fill
                    c.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
                    c.border=Border(left=Side(style="thin", color="808080"), right=Side(style="thin", color="808080"), top=Side(style="thin", color="808080"), bottom=Side(style="thin", color="808080"))
                # Give every cell in the merged header areas a border so the visual box surrounds the logo/title.
                thin=Side(style="thin", color="808080")
                for row in ws.iter_rows(min_row=1,max_row=3,min_col=1,max_col=11):
                    for c in row:
                        c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
                for cell in ws[4]:
                    cell.font=copy(cell.font); cell.font=cell.font.copy(bold=True, color="FFFFFF")
                    cell.fill=PatternFill("solid", fgColor="17365D")
                    cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
                # All data cells are centered/middle and wrapped; long narrative fields get extra width.
                for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=1, max_col=len(preferred)):
                    for cell in row:
                        cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
                widths=[8,20,20,20,20,20,50,50,20,20,20]
                for i,w in enumerate(widths,1):
                    ws.column_dimensions[chr(64+i)].width=w
                ws.freeze_panes="A5"
                ws.auto_filter.ref=f"A4:K{ws.max_row}"
                # Real Excel table for the register rows.
                from openpyxl.worksheet.table import Table, TableStyleInfo
                if ws.max_row >= 4:
                    tab=Table(displayName="HSEObservationRegister", ref=f"A4:K{ws.max_row}")
                    style=TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                    tab.tableStyleInfo=style
                    ws.add_table(tab)
                target=Path(path); temp=target.with_name(target.name + ".tmp.xlsx")
                try:
                    wb.save(temp)
                    with zipfile.ZipFile(temp,"r") as z:
                        bad=z.testzip()
                        if bad: raise ValueError(f"Generated Excel report is corrupt: {bad}")
                    os.replace(temp,target)
                finally: temp.unlink(missing_ok=True)
                self.show_export_success(path); return
            wb=Workbook(); ws=wb.active; ws.title=table_name[:31]; ws.append(columns)
            for row in data: ws.append([xml_safe(v) for v in row])
            for cell in ws[1]:
                cell.font=copy(cell.font); cell.font=cell.font.copy(bold=True)
            ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
            if "Attachments" in columns:
                col=columns.index("Attachments")+1
                for r in range(2, ws.max_row+1):
                    ws.cell(r,col).alignment=Alignment(wrap_text=True, vertical="top")
                    # Make the attachment cell readable; the dedicated Attachments sheet below contains file-level hyperlinks.
                    ws.cell(r,col).hyperlink = f"#'Attachments'!A1"
            for column_cells in ws.columns:
                letter=column_cells[0].column_letter
                max_len=min(55,max(len(xml_safe(c.value)) if c.value is not None else 0 for c in column_cells)+2)
                ws.column_dimensions[letter].width=max(12,max_len)

            # Include a dedicated attachment register in the Excel report so evidence is
            # not lost when the main register has many columns.
            attach_table=attachment_export_map(table_name)
            if attach_table:
                ids=[r["id"] for r in db.fetchall(f"SELECT id FROM {table_name} ORDER BY id DESC") if not record_ids or r["id"] in record_ids]
                aws=wb.create_sheet("Attachments")
                aws.append(["Record ID","Record Number","File Name","Attachment Type","Stored File Path"])
                for rid in ids:
                    rec=db.fetchone(f"SELECT number FROM {table_name} WHERE id=?",(rid,))
                    for a in attachment_rows(attach_table,rid):
                        fpath=safe(a["file_path"])
                        aws.append([rid, safe(rec["number"]) if rec else "", Path(fpath).name, xml_safe(a["attachment_type"]), fpath])
                        cell=aws.cell(aws.max_row,5)
                        if fpath and Path(fpath).exists():
                            cell.hyperlink=Path(fpath).as_uri()
                            cell.style="Hyperlink"
                aws.freeze_panes="A2"; aws.auto_filter.ref=aws.dimensions
                for column_cells in aws.columns:
                    letter=column_cells[0].column_letter
                    max_len=min(80,max(len(xml_safe(c.value)) if c.value is not None else 0 for c in column_cells)+2)
                    aws.column_dimensions[letter].width=max(12,max_len)
            target=Path(path)
            temp=target.with_name(target.name + ".tmp.xlsx")
            try:
                wb.save(temp)
                with zipfile.ZipFile(temp, "r") as z:
                    bad=z.testzip()
                    if bad: raise ValueError(f"Generated Excel report is corrupt: {bad}")
                os.replace(temp,target)
            finally:
                temp.unlink(missing_ok=True)
            self.show_export_success(path)
        except Exception as e:
            logging.exception("Excel export failed"); QMessageBox.critical(self,"Export Error",str(e))

    def export_pdf(self, table_name, record_ids=None):
        try:
            columns,data=self.table_data(table_name, record_ids)
            if not columns:
                QMessageBox.information(self,"No Data","There is no data to export."); return
            path=self._export_path(f"{table_name}_report.pdf","Download PDF Report","PDF Files (*.pdf)")
            if not path:return
            styles=getSampleStyleSheet()
            project=db.setting("project_name","")
            if table_name == "observations":
                # Observation PDF header mirrors the fixed Excel report header.
                def draw_page_border(canvas, doc):
                    canvas.saveState()
                    canvas.setStrokeColor(colors.HexColor("#808080"))
                    canvas.setLineWidth(0.8)
                    page_w,page_h=landscape(A4)
                    canvas.rect(12,12,page_w-24,page_h-24,stroke=1,fill=0)
                    canvas.restoreState()
                doc=SimpleDocTemplate(path,pagesize=landscape(A4),rightMargin=22,leftMargin=22,topMargin=92,bottomMargin=24,
                                      onFirstPage=draw_page_border,onLaterPages=draw_page_border)
                story=[]
                report_no=f"{document_prefix()}-OBS-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                logo_path=report_logo_path()
                logo_flow=Paragraph(xml_safe(company_name()),styles["Normal"])
                if logo_path and Path(logo_path).exists():
                    try:
                        logo_flow=RLImage(str(logo_path),width=85,height=55,preserveAspectRatio=True)
                    except Exception:
                        logging.exception("Unable to embed company logo in observation PDF header")
                title_style=styles["Title"]
                title_style.fontSize=15; title_style.leading=17; title_style.alignment=1
                meta_style=styles["Normal"]
                meta_style.fontSize=7.5; meta_style.leading=9
                header_right=Table([
                    [Paragraph(f"<b>Report No.:</b> {xml_safe(report_no)}",meta_style)],
                    [Paragraph(f"<b>Document No.:</b> {xml_safe(document_prefix())}-OBS",meta_style)],
                    [Paragraph(f"<b>Project:</b> {xml_safe(project)}",meta_style)],
                ], colWidths=[185], rowHeights=[20,20,20])
                header_right.setStyle(TableStyle([
                    ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#808080")),
                    ("INNERGRID",(0,0),(-1,-1),0.3,colors.HexColor("#BFBFBF")),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                    ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),
                ]))
                header=Table([[logo_flow,Paragraph(f"<b>HSE OBSERVATION REPORT</b><br/>{xml_safe(company_name())}",title_style),header_right]],
                             colWidths=[130,455,185], rowHeights=[60])
                header.setStyle(TableStyle([
                    ("BOX",(0,0),(-1,-1),0.7,colors.HexColor("#808080")),
                    ("INNERGRID",(0,0),(-1,-1),0.4,colors.HexColor("#BFBFBF")),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                    ("ALIGN",(0,0),(-1,-1),"CENTER"),
                    ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),
                ]))
                story.append(header)
                story.append(Spacer(1,8))
            else:
                def draw_page_border(canvas, doc):
                    canvas.saveState(); canvas.setStrokeColor(colors.HexColor("#808080")); canvas.setLineWidth(0.8); canvas.rect(12,12,landscape(A4)[0]-24,landscape(A4)[1]-24,stroke=1,fill=0); canvas.restoreState()
                doc=SimpleDocTemplate(path,pagesize=landscape(A4),rightMargin=22,leftMargin=22,topMargin=90,bottomMargin=24,onFirstPage=draw_page_border,onLaterPages=draw_page_border)
                story=[]
                report_no=f"{document_prefix()}-{table_name[:3].upper()}-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                logo_path=report_logo_path()
                logo_flow=Paragraph(xml_safe(company_name()),styles["Normal"])
                if logo_path and Path(logo_path).exists():
                    try: logo_flow=RLImage(str(logo_path),width=85,height=55,preserveAspectRatio=True)
                    except Exception: pass
                ms=styles["Normal"]; ms.fontSize=7.5; ms.leading=9
                meta=Table([[Paragraph(f"<b>Report No.:</b> {xml_safe(report_no)}",ms)],[Paragraph(f"<b>Document No.:</b> {xml_safe(document_prefix())}-{table_name[:3].upper()}",ms)],[Paragraph(f"<b>Project:</b> {xml_safe(project)}",ms)]],colWidths=[185],rowHeights=[20,20,20])
                meta.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#808080")),("INNERGRID",(0,0),(-1,-1),0.3,colors.HexColor("#BFBFBF")),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
                head=Table([[logo_flow,Paragraph(f"<b>{xml_safe(table_name.title())} REPORT</b><br/>{xml_safe(company_name())}",styles["Title"]),meta]],colWidths=[130,455,185],rowHeights=[60])
                head.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.7,colors.HexColor("#808080")),("INNERGRID",(0,0),(-1,-1),0.4,colors.HexColor("#BFBFBF")),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(-1,-1),"CENTER")]))
                story += [head,Spacer(1,8)]
            if table_name == "observations":
                preferred = ["S.No.", "Observation Date", "Location", "Observation Type", "Category", "Observation Detail", "Action Taken", "Responsible", "Evidence", "Status"]
                # Build a visible evidence cell using real images, not hyperlinks.
                obs_rows=db.fetchall("SELECT * FROM observations ORDER BY id DESC")
                if record_ids:
                    obs_rows=[r for r in obs_rows if r["id"] in record_ids]
                pdf_data=[preferred]
                for serial,r in enumerate(obs_rows,1):
                    ps=styles["Normal"]
                    ps.fontSize=5.2; ps.leading=6
                    evidence=[]
                    for a in attachment_rows("observation_attachments",r["id"]):
                        fp=Path(safe(a["file_path"]))
                        if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp"}:
                            try:
                                im=RLImage(str(fp),width=38,height=30,preserveAspectRatio=True)
                                evidence.append(im)
                            except Exception:
                                logging.exception("Unable to embed observation image in PDF")
                        elif fp.exists():
                            evidence.append(Paragraph(xml_safe(fp.name), ps))
                    if not evidence:
                        evidence=[Paragraph("None",ps)]
                    wrap=lambda v: Paragraph(xml_safe(v).replace("\n","<br/>"), ps)
                    pdf_data.append([wrap(serial),wrap(r["obs_date"]),wrap(r["location"]),wrap(r["obs_type"]),wrap(r["category"]),wrap(r["observation"]),wrap(r["corrective_action"]),wrap(r["responsible"]),evidence,wrap(r["status"])])
                header_style=styles["Normal"]
                header_style.fontSize=5.5; header_style.leading=6
                pdf_data[0]=[Paragraph(x,header_style) for x in preferred]
                col_widths=[0.32*72,0.66*72,0.88*72,0.78*72,0.78*72,1.55*72,1.35*72,0.78*72,1.62*72,0.58*72]
                t=Table(pdf_data,repeatRows=1,colWidths=col_widths)
            else:
                if table_name == "incidents" and "Attachments" in columns:
                    display_indices=list(range(9))+[columns.index("Attachments")]
                    pdf_data=[[xml_safe(columns[i]) for i in display_indices]]
                    incident_rows=db.fetchall("SELECT * FROM incidents ORDER BY id DESC")
                    if record_ids:
                        wanted=set(record_ids); incident_rows=[r for r in incident_rows if r["id"] in wanted]
                    cell_style=styles["Normal"]; cell_style.fontSize=5.5; cell_style.leading=6
                    for r in incident_rows[:500]:
                        vals=[xml_safe(r[columns[i]] if columns[i] in r.keys() else "") for i in display_indices[:-1]]
                        evidence=[]
                        for a in attachment_rows("incident_attachments",r["id"]):
                            fp=Path(safe(a["file_path"]))
                            remark=safe(a["remark"]) if "remark" in a.keys() else ""
                            if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                                try:
                                    evidence.append(RLImage(str(fp),width=65,height=48,preserveAspectRatio=True))
                                    evidence.append(Paragraph(xml_safe(fp.name + (f" | {remark}" if remark else "")),cell_style))
                                except Exception:
                                    evidence.append(Paragraph(xml_safe(fp.name + (f" | {remark}" if remark else "")),cell_style))
                            else:
                                evidence.append(Paragraph(xml_safe(fp.name + (f" | {remark}" if remark else "")),cell_style))
                        if not evidence: evidence=[Paragraph("No attachment",cell_style)]
                        pdf_data.append([Paragraph(v.replace("\n","<br/>"),cell_style) for v in vals]+[evidence])
                    t=Table(pdf_data,repeatRows=1,colWidths=[55,55,65,65,65,90,75,65,75,125])
                elif len(columns) > 10 and "Attachments" in columns:
                    display_indices=list(range(9))+[columns.index("Attachments")]
                    pdf_data=[[xml_safe(columns[i]) for i in display_indices]]+[[xml_safe(r[i]) for i in display_indices] for r in data[:500]]
                    t=Table(pdf_data,repeatRows=1)
                else:
                    display_indices=list(range(min(len(columns),10)))
                    pdf_data=[[xml_safe(columns[i]) for i in display_indices]]+[[xml_safe(r[i]) for i in display_indices] for r in data[:500]]
                    t=Table(pdf_data,repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("GRID",(0,0),(-1,-1),0.4,colors.grey),("FONTSIZE",(0,0),(-1,-1),6),("VALIGN",(0,0),(-1,-1),"TOP"),
                ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
                ("TOPPADDING",(0,0),(-1,-1),1.5),("BOTTOMPADDING",(0,0),(-1,-1),1.5)]))
            story.append(t)
            footer=db.setting("report_footer","")
            if footer: story.extend([Spacer(1,10),Paragraph(xml_safe(footer),styles["Normal"])])
            doc.build(story); self.show_export_success(path)
        except Exception as e:
            logging.exception("PDF export failed"); QMessageBox.critical(self,"Export Error",str(e))

    def export_docx(self, table_name, record_ids=None):
        try:
            columns,data=self.table_data(table_name, record_ids)
            if not columns:
                QMessageBox.information(self,"No Data","There is no data to export."); return
            path=self._export_path(f"{table_name}_report.docx","Download Word Report","Word Documents (*.docx)")
            if not path:return
            generate_table_docx(table_name,path,record_ids=record_ids)
            with zipfile.ZipFile(path,"r") as z: z.testzip()
            self.show_export_success(path)
        except Exception as e:
            logging.exception("Word report failed"); QMessageBox.critical(self,"Export Error",str(e))

    def export_table(self, table, filename):
        self._export_table_widget_csv(table, filename)

    def _export_table_widget_csv(self, table, filename):
        path,_=QFileDialog.getSaveFileName(self,"Export CSV",f"{filename}.csv","CSV (*.csv)")
        if not path:return
        with open(path,"w",newline="",encoding="utf-8-sig") as f:
            writer=csv.writer(f)
            writer.writerow([table.horizontalHeaderItem(c).text() for c in range(table.columnCount())])
            for r in range(table.rowCount()):
                writer.writerow([table.item(r,c).text() if table.item(r,c) else "" for c in range(table.columnCount())])
        self.show_export_success(path)

    # ========================================================
    # MASTER DATA
    # ========================================================


    def master_data(self):
        w,layout=self.page("Master Data")
        layout.addWidget(QLabel("Manage employees, companies, projects, categories and company profile."))
        profile=QPushButton("Company Profile / Logo / Document Number")
        profile.clicked.connect(lambda:self.settings_page())
        layout.addWidget(profile)
        for title,table_name in [
            ("Employees","employees"),("Companies","companies"),
            ("Projects / Locations","projects"),("Observation Categories","categories")
        ]:
            button=QPushButton(title)
            button.clicked.connect(lambda checked=False,t=table_name,n=title:self.master_editor(t,n))
            layout.addWidget(button)


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
        w,layout=self.page("Settings")
        form=QFormLayout()
        company=QLineEdit(db.setting("company_name",""))
        project=QLineEdit(db.setting("project_name",""))
        location=QLineEdit(db.setting("default_location",""))
        observer=QLineEdit(db.setting("default_observer",""))
        footer=QLineEdit(db.setting("report_footer",""))
        prefix=QLineEdit(db.setting("document_prefix","HSE"))
        logo_path=QLineEdit(db.setting("company_logo",""))
        logo_path.setReadOnly(True)
        logo_button=QPushButton("Choose Company Logo")
        logo_preview=QLabel()
        logo_preview.setMinimumHeight(80)
        logo_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)

        form.addRow("Company Name",company)
        form.addRow("Project Name",project)
        form.addRow("Default Location",location)
        form.addRow("Default Observer",observer)
        form.addRow("Document Number / Prefix",prefix)
        form.addRow("Report Footer",footer)
        logo_row=QHBoxLayout(); logo_row.addWidget(logo_path); logo_row.addWidget(logo_button)
        form.addRow("Company Logo",logo_row)
        layout.addLayout(form); layout.addWidget(logo_preview)

        def choose_logo():
            path,_=QFileDialog.getOpenFileName(self,"Select Company Logo","","Images (*.png *.jpg *.jpeg *.bmp)")
            if not path:return
            try:
                src=Path(path); dest=ATTACH_DIR / f"company_logo{src.suffix.lower()}"
            except Exception:
                dest=ATTACH_DIR / f"company_logo{Path(path).suffix.lower()}"
            try:
                for old in ATTACH_DIR.glob("company_logo.*"):
                    if old != dest: old.unlink(missing_ok=True)
                shutil.copy2(path,dest)
                logo_path.setText(str(dest))
                pix=QPixmap(str(dest))
                logo_preview.setPixmap(pix.scaled(180,80,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))
            except Exception as e:
                QMessageBox.critical(self,"Logo Error",str(e))

        logo_button.clicked.connect(choose_logo)
        existing_logo=db.setting("company_logo","")
        if existing_logo and Path(existing_logo).exists():
            pix=QPixmap(existing_logo)
            logo_preview.setPixmap(pix.scaled(180,80,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))

        save=QPushButton("Save Settings")
        save.setMinimumHeight(44)
        save.setToolTip("Save company, project and logo settings. New values will be used by all reports.")
        layout.addWidget(save)
        save.clicked.connect(lambda:self.save_settings(company,project,location,observer,footer,prefix,logo_path))
        backup=QPushButton("Backup Complete Application Data"); layout.addWidget(backup); backup.clicked.connect(self.backup)
        restore=QPushButton("Restore Application Backup"); layout.addWidget(restore); restore.clicked.connect(self.restore)
        about=QPushButton("About"); layout.addWidget(about)
        about.clicked.connect(lambda:QMessageBox.information(self,"About",f"{APP_NAME}\nVersion {APP_VERSION}\n\nOffline HSE Management System"))

    def save_settings(self, company, project, location, observer, footer, prefix=None, logo_path=None):
        # Persist all report branding settings in the application database.
        # The logo is copied into the application's attachment directory so the
        # saved report branding does not depend on the user's original file path.
        db.set_setting("company_name", company.text().strip())
        db.set_setting("project_name", project.text().strip())
        db.set_setting("default_location", location.text().strip())
        db.set_setting("default_observer", observer.text().strip())
        db.set_setting("report_footer", footer.text().strip())
        if prefix is not None:
            db.set_setting("document_prefix", prefix.text().strip() or "HSE")

        if logo_path is not None:
            selected = logo_path.text().strip()
            if selected:
                source = Path(selected)
                if source.exists() and source.is_file():
                    try:
                        ATTACH_DIR.mkdir(parents=True, exist_ok=True)
                        suffix = source.suffix.lower() or ".png"
                        destination = ATTACH_DIR / f"company_logo{suffix}"
                        # Remove older saved logo formats so reports always use the
                        # current logo selected in Settings.
                        for old in ATTACH_DIR.glob("company_logo.*"):
                            if old != destination:
                                old.unlink(missing_ok=True)
                        if source.resolve() != destination.resolve():
                            shutil.copy2(source, destination)
                        selected = str(destination)
                    except Exception as e:
                        logging.exception("Unable to persist company logo")
                        QMessageBox.critical(self, "Logo Error", f"Unable to save the company logo.\n\n{e}")
                        return
                elif not source.exists():
                    QMessageBox.warning(self, "Logo Not Found", "The selected company logo could not be found. The previous saved logo will be kept.")
                    selected = db.setting("company_logo", "")
            db.set_setting("company_logo", str(Path(selected).resolve()) if selected and Path(selected).exists() else selected)

        # All report generators read these settings from the database at export
        # time, so the next Word/Excel/PDF/CSV report immediately uses the new
        # company name and logo without changing any other module.
        QMessageBox.information(self, "Saved", "Settings saved successfully. New company name and logo will be used in all new reports.")
        self.dashboard()


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
