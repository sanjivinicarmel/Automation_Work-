"""
Work Reporter - Single-user Daily Notes & Automated Weekly Work Reporting Web App
Features:
1. Daily notes logging (Mon-Fri) with Excel report generation via Gemini AI or fallback.
2. Semester recurring timetable storage in database with Gemini Vision auto-extraction.
3. Daily schedule sent ONE DAY BEFORE at night (e.g. Sunday night for Monday).
4. Human input area to view, add additional classes, or add special notes for tomorrow before sending.
5. Dynamic settings reloading so changes to MANAGER_EMAIL or .env take effect immediately.
6. 100% free-tier architecture with /cron/tick external scheduler.
"""

import os
import re
import json
import base64
import secrets
import smtplib
import zoneinfo
from email.message import EmailMessage
from datetime import datetime, date, timedelta, time, timezone
from functools import wraps
import io

from flask import (
    Flask, render_template_string, request, redirect, url_for,
    session, flash, send_file, Response, jsonify, abort
)
from flask_sqlalchemy import SQLAlchemy
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from dotenv import load_dotenv

# Ensure .env is loaded
load_dotenv(override=True)

# ==============================================================================
# Dynamic Configuration Helpers
# ==============================================================================
def get_config(key, default=""):
    """
    Dynamically loads configuration from .env so any changes made to .env
    (e.g. changing MANAGER_EMAIL or APP_PASSWORD) take effect immediately
    without requiring a server restart.
    """
    load_dotenv(override=True)
    val = os.getenv(key, default)
    return str(val).strip() if val is not None else default


def update_env_file(key_values):
    """
    Safely updates key-value pairs in the .env file while preserving comments.
    """
    env_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), ".env")
    lines = []
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

    updated_keys = set()
    new_lines = []
    for line in lines:
        stripped = line.strip()
        if stripped and not stripped.startswith("#") and "=" in stripped:
            k = stripped.split("=", 1)[0].strip()
            if k in key_values:
                new_lines.append(f"{k}={key_values[k]}\n")
                updated_keys.add(k)
                continue
        new_lines.append(line)

    for k, v in key_values.items():
        if k not in updated_keys:
            new_lines.append(f"{k}={v}\n")

    with open(env_path, "w", encoding="utf-8") as f:
        f.writelines(new_lines)

    load_dotenv(override=True)

# ==============================================================================
# Flask Application Setup
# ==============================================================================
app = Flask(__name__)
app.config["SECRET_KEY"] = get_config("SECRET_KEY", "work-reporter-secret-key-32")

# Database: read DATABASE_URL (Neon Postgres) or fallback to local SQLite
DATABASE_URL = get_config("DATABASE_URL", "")
if DATABASE_URL:
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg2://", 1)
    elif DATABASE_URL.startswith("postgresql://") and not DATABASE_URL.startswith("postgresql+"):
        DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg2://", 1)
    app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
else:
    basedir = os.path.abspath(os.path.dirname(__file__))
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{os.path.join(basedir, 'work_reporter.db')}"

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True}
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024  # 5 MB upload limit
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

db = SQLAlchemy(app)

# ==============================================================================
# Database Models
# ==============================================================================
class DailyNote(db.Model):
    """Stores daily plain-English work notes for a specific date."""
    __tablename__ = "daily_notes"
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, unique=True, nullable=False, index=True)
    content = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class TimetableSlot(db.Model):
    """
    Recurring weekly semester class timetable.
    day_of_week: 0=Monday, 1=Tuesday, 2=Wednesday, 3=Thursday, 4=Friday, 5=Saturday
    """
    __tablename__ = "timetable_slots"
    id = db.Column(db.Integer, primary_key=True)
    day_of_week = db.Column(db.Integer, nullable=False, index=True)
    start_time = db.Column(db.String(20), nullable=False)   # e.g. "09:00 AM"
    end_time = db.Column(db.String(20), nullable=False)     # e.g. "10:00 AM"
    subject = db.Column(db.String(150), nullable=False)    # e.g. "Cloud Computing"
    room = db.Column(db.String(80), nullable=True)         # e.g. "Room 204"
    faculty = db.Column(db.String(100), nullable=True)     # e.g. "Prof. Sharma"
    sort_order = db.Column(db.Integer, default=0)


class ScheduleOverride(db.Model):
    """
    Human input area: allows user to add additional classes, substitute classes,
    or special announcements for a specific date.
    """
    __tablename__ = "schedule_overrides"
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False, index=True)
    action = db.Column(db.String(20), default="add")       # 'add' or 'note'
    start_time = db.Column(db.String(20), nullable=True)   # e.g. "03:00 PM"
    end_time = db.Column(db.String(20), nullable=True)     # e.g. "04:30 PM"
    subject = db.Column(db.String(150), nullable=False)    # e.g. "Special Guest Lecture"
    room = db.Column(db.String(80), nullable=True)         # e.g. "Auditorium"
    faculty = db.Column(db.String(100), nullable=True)     # e.g. "Guest Speaker"
    notes = db.Column(db.Text, nullable=True)              # e.g. "Bring lab journal"
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class TimetableImage(db.Model):
    """Stores the latest timetable screenshot or file as a blob in the database."""
    __tablename__ = "timetable_images"
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    mimetype = db.Column(db.String(150), nullable=False)
    data = db.Column(db.LargeBinary, nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)


class Sent(db.Model):
    """
    Audit log of dispatched automated emails to avoid duplicates on the same day.
    kind: 'weekly_report', 'weekly_report_warning', 'daily_schedule', 'daily_schedule_reminder'
    """
    __tablename__ = "sent"
    id = db.Column(db.Integer, primary_key=True)
    kind = db.Column(db.String(50), nullable=False, index=True)
    target_date = db.Column(db.Date, nullable=False, index=True)
    recipient = db.Column(db.String(255), nullable=False)
    sent_at = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), nullable=False, default="success")
    details = db.Column(db.Text, nullable=True)


with app.app_context():
    db.create_all()

# ==============================================================================
# Timezones & Date Utilities
# ==============================================================================
def get_user_tz():
    """Returns timezone object for configured TIMEZONE, with multi-stage fallback."""
    tz_name = get_config("TIMEZONE", "Asia/Kolkata")
    try:
        return zoneinfo.ZoneInfo(tz_name)
    except Exception:
        try:
            return zoneinfo.ZoneInfo("Asia/Kolkata")
        except Exception:
            return timezone(timedelta(hours=5, minutes=30))


def get_now_user():
    """Returns current datetime in user's configured timezone."""
    return datetime.now(get_user_tz())


def parse_time_str(time_str):
    """Parses 'HH:MM' string into datetime.time object."""
    try:
        parts = time_str.strip().split(":")
        return time(int(parts[0]), int(parts[1]))
    except Exception:
        return time(21, 0)


def get_week_range(ref_date=None, include_weekend=True):
    """
    Given a reference date (defaults to today in user's timezone),
    returns (monday_date, end_date, list_of_dates).
    If include_weekend=True, returns all 7 days (Monday to Sunday) so weekend
    dates (Saturday & Sunday) and today's date are always displayed and editable.
    If include_weekend=False, returns 5 days (Monday to Friday).
    """
    if ref_date is None:
        ref_date = get_now_user().date()
    monday = ref_date - timedelta(days=ref_date.weekday())
    if include_weekend:
        sunday = monday + timedelta(days=6)
        all_days = [monday + timedelta(days=i) for i in range(7)]
        return monday, sunday, all_days
    else:
        friday = monday + timedelta(days=4)
        weekdays = [monday + timedelta(days=i) for i in range(5)]
        return monday, friday, weekdays


def get_weekday_name(day_index):
    """Converts 0..6 to weekday name."""
    names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    return names[day_index % 7]

# ==============================================================================
# Authentication & CSRF
# ==============================================================================
def login_required(f):
    """Decorator to require session-based login for web routes."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login", next=request.url))
        return f(*args, **kwargs)
    return decorated_function


def get_csrf_token():
    """Generates or retrieves CSRF token from current session."""
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_hex(20)
    return session["csrf_token"]


def validate_csrf():
    """Validates CSRF token on POST requests using constant-time comparison."""
    token = request.form.get("csrf_token", "")
    expected = session.get("csrf_token", "")
    if not token or not expected or not secrets.compare_digest(token, expected):
        abort(400, description="Invalid or missing CSRF token.")

# ==============================================================================
# SMTP Email Dispatcher
# ==============================================================================
def send_email(to_email, subject, body_html, body_text=None, cc_email=None,
               attachment_bytes=None, attachment_filename=None, attachment_mimetype=None):
    """
    Sends email using SMTP with STARTTLS (or SSL on port 465).
    Safely cleans whitespace from Gmail App Passwords and falls back gracefully.
    """
    smtp_host = get_config("SMTP_HOST", "smtp.gmail.com")
    try:
        smtp_port = int(get_config("SMTP_PORT", "587"))
    except ValueError:
        smtp_port = 587
    smtp_user = get_config("SMTP_USER", "")
    # Remove spaces from Google App Password (e.g. 'ngaq emor azwk gkxy' -> 'ngaqemorazwkgkxy')
    smtp_pass = get_config("SMTP_PASS", "").replace(" ", "")

    # Clean SMTP_FROM
    smtp_from = get_config("SMTP_FROM", "")
    if not smtp_from or "no-reply@example.com" in smtp_from:
        if smtp_user:
            smtp_from = f'"Work Reporter" <{smtp_user}>'
        else:
            smtp_from = "Work Reporter <no-reply@example.com>"

    owner_email = get_config("OWNER_EMAIL", "")

    if not smtp_host or not smtp_user or not smtp_pass:
        raise ValueError("SMTP configuration incomplete: SMTP_HOST, SMTP_USER, or SMTP_PASS missing.")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = smtp_from
    msg["To"] = to_email
    if cc_email:
        msg["Cc"] = cc_email
    if owner_email:
        msg["Reply-To"] = owner_email

    if not body_text:
        body_text = re.sub(r'<[^>]+>', '', body_html)

    msg.set_content(body_text)
    msg.add_alternative(body_html, subtype="html")

    if attachment_bytes and attachment_filename:
        mimetype = attachment_mimetype or "application/octet-stream"
        maintype, subtype = mimetype.split("/", 1) if "/" in mimetype else ("application", "octet-stream")
        msg.add_attachment(
            attachment_bytes,
            maintype=maintype,
            subtype=subtype,
            filename=attachment_filename
        )

    # Recipients
    recipients = [to_email]
    if cc_email:
        for c in cc_email.split(","):
            c_clean = c.strip()
            if c_clean and c_clean not in recipients:
                recipients.append(c_clean)

    if smtp_port == 465:
        with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30) as server:
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)
    else:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)

    return True

# ==============================================================================
# Semester Timetable & Daily Schedule Logic
# ==============================================================================
def get_schedule_for_date(target_date):
    """
    Returns the complete schedule for target_date:
    1. Recurring semester classes for that day of week (TimetableSlot)
    2. Any specific human input overrides / additions (ScheduleOverride)
    """
    weekday = target_date.weekday()
    regular_slots = TimetableSlot.query.filter_by(day_of_week=weekday).order_by(TimetableSlot.sort_order, TimetableSlot.start_time).all()
    overrides = ScheduleOverride.query.filter_by(date=target_date).order_by(ScheduleOverride.start_time).all()

    additional_classes = [o for o in overrides if o.action == "add"]
    special_notes = [o for o in overrides if o.action == "note" or (o.notes and o.action != "add")]

    has_classes = bool(regular_slots or additional_classes)

    return {
        "date": target_date,
        "weekday": weekday,
        "weekday_name": get_weekday_name(weekday),
        "regular_classes": regular_slots,
        "additional_classes": additional_classes,
        "special_notes": special_notes,
        "has_classes": has_classes
    }


def build_schedule_email_text(schedule_data):
    """
    Builds clean, simple plain-text message for manager:
    Date, Class, and Timings.
    """
    target_date = schedule_data["date"]
    date_label = target_date.strftime("%A, %d %b %Y")
    regular_slots = schedule_data["regular_classes"]
    additional_classes = schedule_data["additional_classes"]
    special_notes = schedule_data["special_notes"]
    has_classes = schedule_data["has_classes"]

    lines = [f"Date: {date_label}", ""]
    if not has_classes:
        lines.append("No classes scheduled for tomorrow.")
    else:
        lines.append("Classes & Timings:")
        for s in regular_slots:
            room_info = f" [{s.room}]" if s.room else ""
            lines.append(f"• {s.start_time} – {s.end_time}: {s.subject}{room_info}")
        for a in additional_classes:
            room_info = f" [{a.room}]" if a.room else ""
            note_info = f" (Note: {a.notes})" if a.notes else ""
            lines.append(f"• {a.start_time} – {a.end_time}: {a.subject} [EXTRA CLASS]{room_info}{note_info}")

    if special_notes:
        lines.append("")
        lines.append("Announcements:")
        for n in special_notes:
            lines.append(f"• {n.subject}: {n.notes or ''}")

    lines.append("")
    lines.append("— Sent via Work Reporter")
    return "\n".join(lines)


def build_schedule_email_html(schedule_data):
    """
    Builds clean, simple HTML message for manager:
    Date, Class, and Timings.
    """
    target_date = schedule_data["date"]
    date_label = target_date.strftime("%A, %d %b %Y")
    regular_slots = schedule_data["regular_classes"]
    additional_classes = schedule_data["additional_classes"]
    special_notes = schedule_data["special_notes"]
    has_classes = schedule_data["has_classes"]

    if not has_classes:
        content_html = """
        <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 24px; text-align: center; margin: 18px 0;">
            <div style="font-size: 28px; margin-bottom: 8px;">🏖️</div>
            <div style="font-size: 16px; font-weight: bold; color: #1e3a8a; margin-bottom: 4px;">No classes scheduled for tomorrow.</div>
            <div style="font-size: 13px; color: #64748b;">Enjoy your day!</div>
        </div>
        """
    else:
        items_html = []
        for s in regular_slots:
            room_pill = f'<span style="background: #e0e7ff; color: #3730a3; padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-left: 8px;">{s.room}</span>' if s.room else ''
            items_html.append(f"""
            <div style="display: flex; align-items: flex-start; justify-content: space-between; padding: 12px 14px; border-bottom: 1px solid #f1f5f9; background: #ffffff;">
                <div style="font-weight: 700; color: #1e3a8a; font-size: 14px; min-width: 170px;">🕒 {s.start_time} – {s.end_time}</div>
                <div style="flex: 1; margin-left: 14px; color: #0f172a; font-size: 14px; font-weight: 600;">
                    {s.subject} {room_pill}
                </div>
            </div>
            """)
        for a in additional_classes:
            room_pill = f'<span style="background: #fef3c7; color: #92400e; padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-left: 8px;">{a.room}</span>' if a.room else ''
            note_line = f'<div style="font-size: 12px; color: #b45309; margin-top: 4px;">Note: {a.notes}</div>' if a.notes else ''
            items_html.append(f"""
            <div style="display: flex; align-items: flex-start; justify-content: space-between; padding: 12px 14px; border-bottom: 1px solid #fde68a; background: #fffbeb;">
                <div style="font-weight: 700; color: #b45309; font-size: 14px; min-width: 170px;">
                    ⭐ {a.start_time} – {a.end_time}
                    <div style="font-size: 10px; background: #f59e0b; color: white; display: inline-block; padding: 1px 5px; border-radius: 3px; margin-top: 2px;">ADDITIONAL CLASS</div>
                </div>
                <div style="flex: 1; margin-left: 14px; color: #92400e; font-size: 14px; font-weight: 600;">
                    {a.subject} {room_pill}
                    {note_line}
                </div>
            </div>
            """)
        class_list = "".join(items_html)
        content_html = f"""
        <div style="border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; margin: 18px 0;">
            <div style="background: #f8fafc; padding: 10px 14px; border-bottom: 1px solid #e2e8f0; font-weight: 700; font-size: 12px; color: #475569; text-transform: uppercase; letter-spacing: 0.05em;">
                Classes &amp; Timings
            </div>
            {class_list}
        </div>
        """

    notes_box = ""
    if special_notes:
        notes_items = "".join([f"<li><strong>{n.subject}:</strong> {n.notes or ''}</li>" for n in special_notes])
        notes_box = f"""
        <div style="margin-top: 14px; padding: 10px 14px; background-color: #eff6ff; border-left: 4px solid #3b82f6; border-radius: 4px;">
            <strong style="color: #1e40af; font-size: 13px;">📌 Special Announcements:</strong>
            <ul style="margin: 4px 0 0 0; padding-left: 20px; color: #1e293b; font-size: 13px;">
                {notes_items}
            </ul>
        </div>
        """

    html = f"""
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.5; color: #1e293b; max-width: 580px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; background: #ffffff;">
        <div style="background-color: #1e3a8a; color: #ffffff; padding: 16px 20px;">
            <div style="font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em; opacity: 0.85;">Daily Class Schedule</div>
            <h2 style="margin: 4px 0 0 0; font-size: 1.2rem; font-weight: 700;">Date: {date_label}</h2>
        </div>

        <div style="padding: 20px;">
            <p style="margin: 0 0 10px 0; font-size: 14px; color: #334155;">
                Hello, here is my scheduled class timing for tomorrow:
            </p>

            {content_html}
            {notes_box}

            <div style="margin-top: 20px; padding-top: 12px; border-top: 1px solid #f1f5f9; font-size: 12px; color: #94a3b8;">
                Sent the day before via Work Reporter
            </div>
        </div>
    </div>
    """
    return html


def dispatch_daily_schedule(target_date=None, recipient_email=None):
    """
    Dispatches tomorrow's schedule email:
    - Sent one day before at 6:00 PM (or manual button).
    - Sends a simple, direct message on: Date, Class, and Timings.
    - If classes exist: lists regular & additional classes.
    - If NO classes exist: sends message stating 'No classes scheduled for tomorrow'.
    - Sends to recipient_email (or MANAGER_EMAIL), CC OWNER_EMAIL.
    - Records in Sent table.
    """
    if target_date is None:
        target_date = get_now_user().date() + timedelta(days=1)

    manager_email = recipient_email or get_config("MANAGER_EMAIL")
    owner_email = get_config("OWNER_EMAIL")

    schedule_data = get_schedule_for_date(target_date)
    date_label = target_date.strftime("%A, %d %b %Y")
    has_classes = schedule_data["has_classes"]

    # Subject line
    if has_classes:
        total_classes = len(schedule_data["regular_classes"]) + len(schedule_data["additional_classes"])
        subject = f"Schedule for {date_label} ({total_classes} {'class' if total_classes == 1 else 'classes'})"
    else:
        subject = f"Schedule for {date_label}: No Classes"

    body_html = build_schedule_email_html(schedule_data)
    body_text = build_schedule_email_text(schedule_data)

    try:
        send_email(
            to_email=manager_email,
            subject=subject,
            body_html=body_html,
            body_text=body_text,
            cc_email=owner_email
        )

        sent_record = Sent(
            kind="daily_schedule",
            target_date=target_date,
            recipient=manager_email,
            status="success",
            details=f"Dispatched schedule for {date_label} (has_classes={has_classes}) to {manager_email} (CC {owner_email})."
        )
        db.session.add(sent_record)
        db.session.commit()

        if has_classes:
            return True, f"Schedule for {date_label} sent successfully to {manager_email} (CC'd {owner_email})."
        else:
            return True, f"Notice ('No classes scheduled') for {date_label} sent successfully to {manager_email} (CC'd {owner_email})."
    except Exception as e:
        return False, f"Failed to send schedule email: {str(e)}"

# ==============================================================================
# Timetable Parsing: Excel (.xlsx, .csv) & Gemini AI Multimodal Vision
# ==============================================================================
def normalize_time_slot(slot_str):
    """
    Normalizes time slot strings like:
    '8:00- 9:00', '12:00-1:00', '1:00-2:00', '10:00 AM - 11:00 AM'
    into clean 12-hour tuples: ('08:00 AM', '09:00 AM')
    """
    if not slot_str:
        return "09:00 AM", "10:00 AM"
    slot_str = str(slot_str).strip()
    parts = re.split(r"[-–to]+", slot_str)
    if len(parts) != 2:
        return slot_str, slot_str

    def parse_part(p, is_end=False, start_h=None):
        p = p.strip()
        m = re.match(r"(\d+)(?::(\d+))?\s*(am|pm)?", p, re.I)
        if not m:
            return p, 0, "AM"
        h = int(m.group(1))
        mins = m.group(2) or "00"
        ampm = m.group(3)
        if ampm:
            ampm = ampm.upper()
        else:
            if 8 <= h < 12:
                ampm = "AM"
            elif h == 12:
                ampm = "PM"
            elif h < 8:
                ampm = "PM"
            else:
                ampm = "AM"
        return f"{h:02d}:{mins} {ampm}", h, ampm

    start_norm, start_h, start_ampm = parse_part(parts[0], is_end=False)
    end_norm, end_h, end_ampm = parse_part(parts[1], is_end=True, start_h=start_h)
    if start_h == 11 and end_h == 12:
        end_norm = "12:00 PM"
    elif start_h == 12 and end_h == 1:
        start_norm = "12:00 PM"
        end_norm = "01:00 PM"
    return start_norm, end_norm


def parse_timetable_excel(file_bytes, filename=None):
    """
    Parses a timetable spreadsheet (.xlsx or .csv) using openpyxl or csv:
    - Robustly finds Day column (even if Column A is blank!)
    - Identifies time header row across any column positions
    - Detects legend table at the bottom (e.g. DMW -> Data Mining and Warehousing)
    - Normalizes start/end times and expands course acronyms
    - Populates TimetableSlot table in database
    """
    import csv

    rows = []
    is_csv = filename and filename.lower().endswith(".csv")
    if is_csv:
        try:
            reader = csv.reader(io.StringIO(file_bytes.decode("utf-8", errors="ignore")))
            rows = list(reader)
        except Exception as e:
            return None, f"Failed to read CSV: {str(e)}"
    else:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
        except Exception as e:
            return None, f"Failed to load Excel file: {str(e)}"

    if not rows:
        return None, "Spreadsheet is empty."

    day_map = {
        "monday": 0, "mon": 0,
        "tuesday": 1, "tue": 1, "tues": 1,
        "wednesday": 2, "wed": 2,
        "thursday": 3, "thu": 3, "thurs": 3,
        "friday": 4, "fri": 4,
        "saturday": 5, "sat": 5,
        "sunday": 6, "sun": 6
    }

    # 1. Robustly find the column that contains day names (e.g. Column A, Column B, etc.)
    day_col_idx = None
    day_rows = []
    for r_idx, row in enumerate(rows):
        if not row or not any(row):
            continue
        for c_idx, cell in enumerate(row):
            if cell is not None:
                clean_val = str(cell).strip().lower()
                if clean_val in day_map:
                    if day_col_idx is None:
                        day_col_idx = c_idx
                    if c_idx == day_col_idx:
                        day_rows.append((day_map[clean_val], row))
                    break

    # 2. Find time slot header row (looks for e.g. "8:00- 9:00", "10:00-11:00")
    time_headers = {}
    for r_idx, row in enumerate(rows):
        time_cells = []
        for c_idx, cell in enumerate(row):
            if cell is not None:
                c_str = str(cell).strip()
                if re.search(r"\d+[:.]\d+\s*[-–to]+\s*\d+[:.]\d+", c_str):
                    time_cells.append((c_idx, c_str))
        if len(time_cells) >= 2:
            for c_idx, c_str in time_cells:
                start_t, end_t = normalize_time_slot(c_str)
                time_headers[c_idx] = (start_t, end_t)
            break

    if not time_headers:
        # Fallback: check each cell in row for time patterns
        for r_idx, row in enumerate(rows):
            for c_idx, cell in enumerate(row):
                if cell is not None and re.search(r"\d+[:.]\d+", str(cell)):
                    start_t, end_t = normalize_time_slot(str(cell))
                    time_headers[c_idx] = (start_t, end_t)
            if len(time_headers) >= 2:
                break

    if not time_headers:
        return None, "Could not detect time slot headers (e.g. 8:00- 9:00, 10:00-11:00) in spreadsheet."

    if not day_rows:
        return None, "Could not detect any day rows (Monday, Tuesday, etc.) in spreadsheet."

    # 3. Detect Legend Table (Acronym -> Full Course Name)
    legend = {}
    for row in rows:
        non_empty = [(c_idx, str(c).strip()) for c_idx, c in enumerate(row) if c is not None and str(c).strip()]
        if len(non_empty) >= 2:
            first_txt = non_empty[0][1]
            second_txt = non_empty[1][1]
            if first_txt.lower() not in day_map and len(first_txt) <= 12 and len(second_txt) > 4:
                if not re.search(r"\d+[:.]\d+", first_txt) and not re.search(r"\d+[:.]\d+", second_txt):
                    legend[first_txt.upper()] = second_txt

    DEFAULT_SUBJECT_ALIASES = {
        "DMW": "Data Mining and Warehousing",
        "GEN AI": "Generative Models and Application",
        "DLMA": "Deep Learning Models and Application",
        "NLP": "Natural Language Processing"
    }

    # 4. Extract class slots
    extracted_slots = []
    count = 0
    for day_idx, row in day_rows:
        for c_idx, (start_t, end_t) in time_headers.items():
            if c_idx < len(row):
                val = row[c_idx]
                if val is not None and str(val).strip():
                    cell_text = str(val).strip()
                    parts = [p.strip() for p in cell_text.split(",") if p.strip()]
                    raw_subj = parts[0]
                    room_sec = ", ".join(parts[1:]) if len(parts) > 1 else ""

                    raw_subj_clean = re.sub(r"\s+", " ", raw_subj).strip()
                    room_sec_clean = re.sub(r"\s+", " ", room_sec).strip()
                    room_sec_clean = re.sub(r"&\s*", "& ", room_sec_clean).strip()
                    room_sec_clean = re.sub(r"\s+", " ", room_sec_clean).strip()

                    expanded_subj = legend.get(raw_subj_clean.upper(), DEFAULT_SUBJECT_ALIASES.get(raw_subj_clean.upper(), raw_subj_clean))
                    if expanded_subj != raw_subj_clean:
                        final_subj = f"{expanded_subj} ({raw_subj_clean})"
                    else:
                        final_subj = raw_subj_clean

                    extracted_slots.append(TimetableSlot(
                        day_of_week=day_idx,
                        start_time=start_t,
                        end_time=end_t,
                        subject=final_subj,
                        room=room_sec_clean,
                        faculty="",
                        sort_order=count
                    ))
                    count += 1

    if not extracted_slots:
        return 0, "No class entries found in the detected day rows."

    # Replace existing slots in database
    TimetableSlot.query.delete()
    for s in extracted_slots:
        db.session.add(s)
    db.session.commit()

    return count, f"Successfully parsed Excel timetable and saved {count} semester class slots!"


def convert_to_download_url(url):
    """
    Converts sharing URLs from Google Sheets, Dropbox, or OneDrive/SharePoint
    into direct download URLs for .xlsx.
    """
    if not url:
        return "", ""
    url = url.strip()

    # Google Sheets: /edit... -> /export?format=xlsx
    gsheet_match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if gsheet_match:
        sheet_id = gsheet_match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx", "google_sheet_timetable.xlsx"

    # Dropbox: dl=0 -> dl=1
    if "dropbox.com" in url:
        return re.sub(r"[?&]dl=0", "?dl=1", url), "dropbox_timetable.xlsx"

    # OneDrive / SharePoint: add download=1
    if any(k in url for k in ["sharepoint.com", "1drv.ms", "onedrive.live.com"]):
        if "download=1" not in url:
            sep = "&" if "?" in url else "?"
            return f"{url}{sep}download=1", "onedrive_timetable.xlsx"

    # Direct link or raw URL
    filename = url.split("/")[-1].split("?")[0] or "cloud_timetable.xlsx"
    return url, filename


def sync_timetable_from_url(cloud_url=None):
    """
    Downloads timetable from Google Sheets / OneDrive / Web Excel URL and parses it.
    """
    url = cloud_url or get_config("TIMETABLE_EXCEL_URL", "").strip()
    if not url:
        return None, "No cloud Excel URL is configured."

    download_url, filename = convert_to_download_url(url)
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) WorkReporter/1.0"}
        resp = requests.get(download_url, headers=headers, timeout=25)
        if resp.status_code != 200:
            return None, f"Failed to download cloud sheet (HTTP {resp.status_code}). Please verify the link is public/shared."

        file_bytes = resp.content
        if len(file_bytes) < 50:
            return None, "Downloaded file is empty or link is invalid."

        # Save to database
        TimetableImage.query.delete()
        new_file = TimetableImage(
            filename=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            data=file_bytes,
            uploaded_at=datetime.utcnow()
        )
        db.session.add(new_file)
        db.session.commit()

        count, msg = parse_timetable_excel(file_bytes, filename=filename)
        return count, f"Cloud sync successful! {msg}"
    except Exception as e:
        return None, f"Cloud sync error: {str(e)}"


def extract_timetable_from_image(image_bytes, mimetype):
    """
    Calls Google Gemini Multimodal Vision API to parse the uploaded semester timetable image.
    Extracts recurring weekly slots (Monday to Sunday, 0..6).
    """
    gemini_key = get_config("GEMINI_API_KEY")
    if not gemini_key:
        return None, "Gemini API key is not configured. Timetable file saved, but AI extraction was skipped."

    model_name = get_config("GEMINI_MODEL", "gemini-1.5-flash")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={gemini_key}"

    prompt = (
        "Extract the recurring weekly semester class schedule/timetable from this document/image.\n"
        "Days can be Monday (0), Tuesday (1), Wednesday (2), Thursday (3), Friday (4), Saturday (5), Sunday (6).\n"
        "Return a JSON array of class slot objects. Each object MUST have exactly these keys:\n"
        "  - \"day_of_week\": integer (0 for Monday, 1 for Tuesday, 2 for Wednesday, 3 for Thursday, 4 for Friday, 5 for Saturday, 6 for Sunday)\n"
        "  - \"start_time\": string (e.g. \"08:00 AM\", \"10:00 AM\", \"12:00 PM\")\n"
        "  - \"end_time\": string (e.g. \"09:00 AM\", \"11:00 AM\", \"01:00 PM\")\n"
        "  - \"subject\": string (course/subject name; if acronyms like DMW, GEN AI, DLMA, NLP are expanded at the bottom legend, write the full name like 'Data Mining and Warehousing (DMW)')\n"
        "  - \"room\": string (room number, section like 'MCA B', or empty string if not shown)\n"
        "  - \"faculty\": string (professor/instructor name, or empty string if not shown)\n\n"
        "CRITICAL: Do NOT invent classes. Return ONLY a valid JSON array. No markdown backticks or commentary."
    )

    b64_img = base64.b64encode(image_bytes).decode("utf-8")
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {
                    "inlineData": {
                        "mimeType": mimetype,
                        "data": b64_img
                    }
                }
            ]
        }],
        "generationConfig": {
            "responseMimeType": "application/json",
            "temperature": 0.1
        }
    }

    try:
        resp = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=35)
        if resp.status_code == 200:
            res_json = resp.json()
            candidates = res_json.get("candidates", [])
            if candidates:
                raw_text = candidates[0].get("content", {}).get("parts", [{}])[0].get("text", "").strip()
                if raw_text.startswith("```"):
                    raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
                    raw_text = re.sub(r"\s*```$", "", raw_text)
                slots = json.loads(raw_text)
                if isinstance(slots, list) and len(slots) > 0:
                    TimetableSlot.query.delete()
                    count = 0
                    for s in slots:
                        if isinstance(s, dict) and "subject" in s and "day_of_week" in s:
                            new_slot = TimetableSlot(
                                day_of_week=int(s["day_of_week"]),
                                start_time=str(s.get("start_time", "09:00 AM")),
                                end_time=str(s.get("end_time", "10:00 AM")),
                                subject=str(s.get("subject", "Class")),
                                room=str(s.get("room", "")),
                                faculty=str(s.get("faculty", "")),
                                sort_order=count
                            )
                            db.session.add(new_slot)
                            count += 1
                    db.session.commit()
                    return count, f"Successfully extracted and saved {count} semester class slots from timetable image!"
        return None, f"Gemini could not parse class slots from image (HTTP {resp.status_code})."
    except Exception as e:
        return None, f"AI extraction error: {str(e)}"

# ==============================================================================
# Weekly Notes & Excel Report Logic
# ==============================================================================
def parse_notes_fallback(notes_by_day):
    """Rule-based parser that splits notes on commas, semicolons, 'and', and newlines."""
    parsed_tasks = []
    category_keywords = {
        "Meeting": ["meeting", "standup", "planning", "retro", "sync", "call", "1:1", "demo", "scrum", "discussion"],
        "Review": ["review", "pr", "pull request", "code review", "feedback", "reviewed"],
        "Documentation": ["doc", "docs", "documentation", "readme", "guide", "paper", "write-up", "report"],
        "DevOps": ["deploy", "deployment", "docker", "release", "ci/cd", "pipeline", "aws", "render", "neon", "infra", "server", "cloud"],
        "Development": ["fix", "fixed", "bug", "implement", "implemented", "build", "refactor", "feature", "code", "coding", "test", "testing", "debug", "api", "ui", "research"]
    }

    for day_name, d_date, text in notes_by_day:
        if not text or not text.strip():
            continue
        raw_items = re.split(r'[\r\n;,]+|\band\b', text, flags=re.IGNORECASE)
        for item in raw_items:
            task_clean = item.strip()
            task_clean = re.sub(r'^[\*\-\•\d+\.\s]+', '', task_clean).strip()
            if not task_clean or len(task_clean) < 2:
                continue

            # Extract hours
            hours = None
            hours_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:hrs?|hours?|h\b)', task_clean, re.IGNORECASE)
            if hours_match:
                try:
                    val = float(hours_match.group(1))
                    hours = int(val) if val.is_integer() else val
                except Exception:
                    hours = None
            else:
                mins_match = re.search(r'(\d+)\s*(?:mins?|minutes?)', task_clean, re.IGNORECASE)
                if mins_match:
                    try:
                        hours = round(int(mins_match.group(1)) / 60.0, 2)
                    except Exception:
                        hours = None

            lower = task_clean.lower()
            detected_category = "General"
            for cat, kws in category_keywords.items():
                if any(kw in lower for kw in kws):
                    detected_category = cat
                    break

            status = "Completed"
            if any(term in lower for term in ["in progress", "wip", "started", "working on", "pending", "halfway"]):
                status = "In Progress"
            elif any(term in lower for term in ["blocked", "stuck", "waiting on", "waiting for"]):
                status = "Blocked"

            display_task = task_clean[0].upper() + task_clean[1:] if len(task_clean) > 0 else task_clean

            parsed_tasks.append({
                "day": day_name,
                "raw_day": day_name,
                "date": d_date.isoformat(),
                "task": display_task,
                "category": detected_category,
                "hours": hours,
                "status": status
            })

    return parsed_tasks


def parse_notes_with_gemini_or_fallback(notes_by_day):
    """Parses notes with Gemini API if key is present, else uses fallback parser."""
    if not notes_by_day:
        return []

    gemini_key = get_config("GEMINI_API_KEY")
    if not gemini_key:
        return parse_notes_fallback(notes_by_day)

    notes_blocks = [f"Day: {d_name} ({d_date.isoformat()}):\n{text.strip()}" for d_name, d_date, text in notes_by_day]
    notes_prompt_text = "\n\n".join(notes_blocks)

    system_prompt = (
        "Organize the employee's daily work notes into a clean JSON list of individual tasks.\n"
        "Each JSON object must have EXACTLY these keys:\n"
        '  - "day": string (e.g. "Monday", "Tuesday", etc.)\n'
        '  - "task": string (clear description)\n'
        '  - "category": string (e.g. "Development", "Meeting", "Review", "DevOps", "Documentation", "Research", or "General")\n'
        '  - "hours": number or null\n'
        '  - "status": string (e.g. "Completed", "In Progress", "Blocked")\n\n'
        "Do NOT invent tasks. Return ONLY valid JSON array.\n\n"
        f"Notes:\n{notes_prompt_text}"
    )

    model_name = get_config("GEMINI_MODEL", "gemini-1.5-flash")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={gemini_key}"
    payload = {
        "contents": [{"parts": [{"text": system_prompt}]}],
        "generationConfig": {"responseMimeType": "application/json", "temperature": 0.1}
    }

    try:
        resp = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=20)
        if resp.status_code == 200:
            candidates = resp.json().get("candidates", [])
            if candidates:
                raw_text = candidates[0].get("content", {}).get("parts", [{}])[0].get("text", "").strip()
                if raw_text.startswith("```"):
                    raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
                    raw_text = re.sub(r"\s*```$", "", raw_text)
                parsed = json.loads(raw_text)
                if isinstance(parsed, list) and len(parsed) > 0:
                    cleaned_tasks = []
                    for item in parsed:
                        if isinstance(item, dict) and "task" in item:
                            cleaned_tasks.append({
                                "day": str(item.get("day", "Unknown")),
                                "task": str(item.get("task", "")),
                                "category": str(item.get("category", "General")),
                                "hours": item.get("hours") if isinstance(item.get("hours"), (int, float)) else None,
                                "status": str(item.get("status", "Completed"))
                            })
                    if cleaned_tasks:
                        return cleaned_tasks
        return parse_notes_fallback(notes_by_day)
    except Exception:
        return parse_notes_fallback(notes_by_day)


def generate_weekly_report_excel(start_date, end_date, tasks_list, weekdays_meta):
    """Builds OpenPyXL workbook with 'Weekly Report' and 'Summary' sheets."""
    wb = openpyxl.Workbook()

    # Sheet 1: Weekly Report
    ws_report = wb.active
    ws_report.title = "Weekly Report"
    ws_report.views.sheetView[0].showGridLines = True

    TITLE_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    HEADER_FILL = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    TOTAL_FILL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    BOLD_FONT = Font(name="Calibri", size=11, bold=True, color="0F172A")
    REGULAR_FONT = Font(name="Calibri", size=11, color="0F172A")
    MUTED_FONT = Font(name="Calibri", size=10, italic=True, color="64748B")

    THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
    GRID_BORDER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
    DOUBLE_BOTTOM_SIDE = Side(border_style="double", color="334155")
    TOTAL_BORDER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=DOUBLE_BOTTOM_SIDE)

    CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
    LEFT_WRAP_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Title Row
    title_text = f"Weekly Work Report: {start_date.strftime('%d %b %Y')} – {end_date.strftime('%d %b %Y')}"
    ws_report.merge_cells("A1:E1")
    ws_report["A1"].value = title_text
    ws_report["A1"].font = TITLE_FONT
    ws_report["A1"].fill = TITLE_FILL
    ws_report["A1"].alignment = CENTER_ALIGN
    ws_report.row_dimensions[1].height = 36

    # Subtitle Row
    ws_report.merge_cells("A2:E2")
    ws_report["A2"].value = f"Generated by Work Reporter • Total tasks: {len(tasks_list)}"
    ws_report["A2"].font = MUTED_FONT
    ws_report["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws_report.row_dimensions[2].height = 20

    # Headers
    headers = ["Day", "Task", "Category", "Hours", "Status"]
    ws_report.row_dimensions[3].height = 26
    for col_idx, col_name in enumerate(headers, start=1):
        cell = ws_report.cell(row=3, column=col_idx, value=col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = GRID_BORDER

    current_row = 4
    if tasks_list:
        for idx, item in enumerate(tasks_list):
            row_fill = ZEBRA_FILL if idx % 2 == 1 else WHITE_FILL

            c_day = ws_report.cell(row=current_row, column=1, value=item.get("day", ""))
            c_day.alignment = CENTER_ALIGN

            c_task = ws_report.cell(row=current_row, column=2, value=item.get("task", ""))
            c_task.alignment = LEFT_WRAP_ALIGN

            c_cat = ws_report.cell(row=current_row, column=3, value=item.get("category", "General"))
            c_cat.alignment = CENTER_ALIGN

            hours_val = item.get("hours")
            c_hours = ws_report.cell(row=current_row, column=4, value=hours_val if hours_val is not None else "-")
            c_hours.alignment = CENTER_ALIGN
            if isinstance(hours_val, (int, float)):
                c_hours.number_format = "#,##0.0"

            c_status = ws_report.cell(row=current_row, column=5, value=item.get("status", "Completed"))
            c_status.alignment = CENTER_ALIGN

            for cell in (c_day, c_task, c_cat, c_hours, c_status):
                cell.font = REGULAR_FONT
                cell.fill = row_fill
                cell.border = GRID_BORDER

            ws_report.row_dimensions[current_row].height = 28
            current_row += 1
    else:
        ws_report.merge_cells("A4:E4")
        empty_cell = ws_report["A4"]
        empty_cell.value = "No work notes recorded for this week."
        empty_cell.font = MUTED_FONT
        empty_cell.alignment = CENTER_ALIGN
        ws_report.row_dimensions[4].height = 30

    ws_report.column_dimensions["A"].width = 18
    ws_report.column_dimensions["B"].width = 54
    ws_report.column_dimensions["C"].width = 20
    ws_report.column_dimensions["D"].width = 14
    ws_report.column_dimensions["E"].width = 18

    # Sheet 2: Summary
    ws_summary = wb.create_sheet(title="Summary")
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary.merge_cells("A1:D1")
    ws_summary["A1"].value = "Weekly Activity Summary"
    ws_summary["A1"].font = TITLE_FONT
    ws_summary["A1"].fill = TITLE_FILL
    ws_summary["A1"].alignment = CENTER_ALIGN
    ws_summary.row_dimensions[1].height = 36

    s_headers = ["Day", "Date", "Task Count", "Total Hours"]
    ws_summary.row_dimensions[3].height = 26
    for col_idx, h_text in enumerate(s_headers, start=1):
        cell = ws_summary.cell(row=3, column=col_idx, value=h_text)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = GRID_BORDER

    s_row = 4
    total_task_count = 0
    total_hours_sum = 0.0

    for idx, d_info in enumerate(weekdays_meta):
        row_fill = ZEBRA_FILL if idx % 2 == 1 else WHITE_FILL

        c_day = ws_summary.cell(row=s_row, column=1, value=d_info["name"])
        c_day.alignment = CENTER_ALIGN

        c_date = ws_summary.cell(row=s_row, column=2, value=d_info["date"].strftime("%d %b %Y"))
        c_date.alignment = CENTER_ALIGN

        c_cnt = ws_summary.cell(row=s_row, column=3, value=d_info["task_count"])
        c_cnt.alignment = CENTER_ALIGN

        hrs = d_info["total_hours"]
        c_hrs = ws_summary.cell(row=s_row, column=4, value=hrs if hrs > 0 else 0.0)
        c_hrs.alignment = CENTER_ALIGN
        c_hrs.number_format = "#,##0.0"

        for cell in (c_day, c_date, c_cnt, c_hrs):
            cell.font = REGULAR_FONT
            cell.fill = row_fill
            cell.border = GRID_BORDER

        total_task_count += d_info["task_count"]
        total_hours_sum += hrs
        ws_summary.row_dimensions[s_row].height = 24
        s_row += 1

    # Total Row
    tot_day = ws_summary.cell(row=s_row, column=1, value="Total")
    tot_day.alignment = CENTER_ALIGN
    tot_date = ws_summary.cell(row=s_row, column=2, value="")
    tot_date.alignment = CENTER_ALIGN
    tot_cnt = ws_summary.cell(row=s_row, column=3, value=total_task_count)
    tot_cnt.alignment = CENTER_ALIGN
    tot_hrs = ws_summary.cell(row=s_row, column=4, value=total_hours_sum)
    tot_hrs.alignment = CENTER_ALIGN
    tot_hrs.number_format = "#,##0.0"

    for cell in (tot_day, tot_date, tot_cnt, tot_hrs):
        cell.font = BOLD_FONT
        cell.fill = TOTAL_FILL
        cell.border = TOTAL_BORDER

    ws_summary.row_dimensions[s_row].height = 26
    ws_summary.column_dimensions["A"].width = 18
    ws_summary.column_dimensions["B"].width = 18
    ws_summary.column_dimensions["C"].width = 18
    ws_summary.column_dimensions["D"].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_week_report_data(ref_date=None):
    """Fetches notes and parses tasks for the full week (Monday through Sunday) of ref_date."""
    monday, sunday, weekdays = get_week_range(ref_date, include_weekend=True)
    notes_map = {n.date: n.content for n in DailyNote.query.filter(DailyNote.date >= monday, DailyNote.date <= sunday).all()}

    notes_by_day = []
    weekdays_meta = []
    has_any_notes = False

    for d in weekdays:
        d_name = d.strftime("%A")
        content = notes_map.get(d, "").strip()
        if content:
            has_any_notes = True
            notes_by_day.append((d_name, d, content))

    parsed_tasks = parse_notes_with_gemini_or_fallback(notes_by_day) if has_any_notes else []

    for d in weekdays:
        d_name = d.strftime("%A")
        d_tasks = [t for t in parsed_tasks if d_name.lower() in t.get("day", "").lower() or t.get("date") == d.isoformat()]
        cnt = len(d_tasks)
        tot_hrs = sum(t["hours"] for t in d_tasks if isinstance(t.get("hours"), (int, float)))
        weekdays_meta.append({
            "name": d_name,
            "date": d,
            "task_count": cnt,
            "total_hours": round(tot_hrs, 2)
        })

    return {
        "monday": monday,
        "friday": sunday,  # Alias so data['friday'] references week end date safely
        "end_date": sunday,
        "weekdays": weekdays,
        "notes_map": notes_map,
        "has_any_notes": has_any_notes,
        "parsed_tasks": parsed_tasks,
        "weekdays_meta": weekdays_meta
    }


def dispatch_weekly_report(target_date=None):
    """Dispatches weekly report email with .xlsx attachment or sends warning to owner if no notes."""
    if target_date is None:
        target_date = get_now_user().date()

    data = build_week_report_data(target_date)
    start_str = data["monday"].strftime("%d %b %Y")
    end_str = data["friday"].strftime("%d %b %Y")
    manager_email = get_config("MANAGER_EMAIL")
    owner_email = get_config("OWNER_EMAIL")

    if not data["has_any_notes"]:
        warning_subject = f"Warning: No Work Notes Recorded ({start_str} – {end_str}) - Report Skipped"
        warning_html = f"""
        <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #1e293b; max-width: 600px;">
            <h2 style="color: #dc2626;">Weekly Report Skipped: No Notes Found</h2>
            <p>Your automated weekly work report for <strong>{start_str} to {end_str}</strong> was scheduled to be sent today.</p>
            <p style="background: #fef2f2; border-left: 4px solid #ef4444; padding: 12px; border-radius: 4px;">
                No notes were found for Monday through Friday. <strong>No email was sent to your manager ({manager_email}).</strong>
            </p>
            <p>Please log in to Work Reporter and add your notes if you wish to send it manually.</p>
        </div>
        """
        try:
            send_email(to_email=owner_email, subject=warning_subject, body_html=warning_html)
            sent_record = Sent(
                kind="weekly_report_warning",
                target_date=target_date,
                recipient=owner_email,
                status="success",
                details="No notes found for week; warning sent to owner."
            )
            db.session.add(sent_record)
            db.session.commit()
            return True, "No notes recorded for this week. Warning sent to owner (manager was not emailed)."
        except Exception as e:
            return False, f"Failed to send warning email: {str(e)}"

    # Notes exist -> Send Excel to manager, CC owner
    try:
        excel_buf = generate_weekly_report_excel(
            data["monday"],
            data["friday"],
            data["parsed_tasks"],
            data["weekdays_meta"]
        )
        report_bytes = excel_buf.getvalue()
        filename = f"Weekly_Report_{data['monday'].isoformat()}_to_{data['friday'].isoformat()}.xlsx"

        subject = f"Weekly Work Report ({start_str} – {end_str})"
        total_tasks = len(data["parsed_tasks"])
        total_hours = sum(d["total_hours"] for d in data["weekdays_meta"])

        body_html = f"""
        <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #1e293b; max-width: 600px;">
            <h2 style="color: #1e3a8a; border-bottom: 2px solid #2563eb; padding-bottom: 8px;">Weekly Work Report</h2>
            <p>Hi,</p>
            <p>Please find attached the weekly work report covering the period from <strong>{start_str}</strong> to <strong>{end_str}</strong>.</p>
            <div style="background: #f1f5f9; padding: 14px 18px; border-radius: 6px; margin: 18px 0;">
                <ul style="margin: 0; padding-left: 18px; color: #334155;">
                    <li><strong>Total Tasks Logged:</strong> {total_tasks}</li>
                    <li><strong>Total Recorded Hours:</strong> {total_hours:.1f} hrs</li>
                    <li><strong>Attachment:</strong> {filename}</li>
                </ul>
            </div>
            <p>The attached Excel workbook includes the itemized task breakdown on the <em>Weekly Report</em> sheet and weekday aggregates on the <em>Summary</em> sheet.</p>
            <p style="color: #64748b; font-size: 13px; margin-top: 24px;">— Sent automatically via Work Reporter</p>
        </div>
        """

        send_email(
            to_email=manager_email,
            subject=subject,
            body_html=body_html,
            cc_email=owner_email,
            attachment_bytes=report_bytes,
            attachment_filename=filename,
            attachment_mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        sent_record = Sent(
            kind="weekly_report",
            target_date=target_date,
            recipient=manager_email,
            status="success",
            details=f"Sent {total_tasks} tasks to manager, CC owner."
        )
        db.session.add(sent_record)
        db.session.commit()
        return True, f"Weekly report ({total_tasks} tasks) sent successfully to {manager_email} (CC'd {owner_email})."
    except Exception as e:
        return False, f"Failed to send weekly report: {str(e)}"

# ==============================================================================
# Cron Tick Endpoint (External Scheduler)
# ==============================================================================
@app.route("/cron/tick", methods=["GET"])
def cron_tick():
    """
    Called every 10 minutes by cron-job.org.
    Sends:
    1. Daily schedule ONE DAY BEFORE at night (e.g. Sunday night for Monday at SCHEDULE_TIME).
    2. Weekly report on REPORT_DAY at REPORT_TIME.
    """
    cron_secret = get_config("CRON_SECRET", "cron-default-secret-token")
    key = request.args.get("key", "")
    if not key or not secrets.compare_digest(key, cron_secret):
        return jsonify({"error": "Forbidden: invalid or missing cron secret key"}), 403

    now = get_now_user()
    today = now.date()
    user_tz = get_user_tz()
    weekday = now.weekday()  # 0=Monday, 6=Sunday

    actions_executed = []

    def is_time_due(sched_time_str):
        target_t = parse_time_str(sched_time_str)
        sched_dt = datetime.combine(today, target_t, tzinfo=user_tz)
        elapsed_seconds = (now - sched_dt).total_seconds()
        # Due if scheduled time arrived and within 2-hour grace window
        return 0 <= elapsed_seconds <= (2 * 3600)

    # 1. Daily Schedule Check:
    # Sent the night before for tomorrow at SCHEDULE_TIME (default 18:00 / 6:00 PM)!
    # Sunday night (6) -> sends Monday (0)
    # Monday night (0) -> sends Tuesday (1)
    # Tuesday night (1) -> sends Wednesday (2)
    # Wednesday night (2) -> sends Thursday (3)
    # Thursday night (3) -> sends Friday (4)
    # Friday night (4) -> sends Saturday (5) if Saturday classes exist (recurring or manual)
    # Saturday night (5) -> sends Sunday (6) if Sunday classes exist (recurring or manual)
    schedule_time_str = get_config("SCHEDULE_TIME", "18:00")
    if is_time_due(schedule_time_str):
        # Auto-sync latest schedule from cloud Excel / Google Sheets URL if configured
        cloud_url = get_config("TIMETABLE_EXCEL_URL", "").strip()
        if cloud_url:
            try:
                sync_timetable_from_url(cloud_url)
            except Exception:
                pass

        tomorrow = today + timedelta(days=1)
        tomorrow_weekday = tomorrow.weekday()
        tomorrow_schedule = get_schedule_for_date(tomorrow)

        # Mon-Fri: always send tomorrow's schedule (classes or "No classes")
        # Sat-Sun: send if classes exist (recurring or manual entry)
        should_send_tomorrow = False
        if tomorrow_weekday < 5:
            should_send_tomorrow = True
        elif tomorrow_schedule["has_classes"]:
            should_send_tomorrow = True

        if should_send_tomorrow:
            already_sent = Sent.query.filter(
                Sent.target_date == tomorrow,
                Sent.kind == "daily_schedule",
                Sent.status == "success"
            ).first()

            if not already_sent:
                success, msg = dispatch_daily_schedule(tomorrow)
                actions_executed.append({
                    "action": "daily_schedule_night_before",
                    "target_date": tomorrow.isoformat(),
                    "success": success,
                    "message": msg
                })

    # 2. Weekly Report Check: REPORT_DAY at REPORT_TIME
    try:
        report_day = int(get_config("REPORT_DAY", "4"))
    except ValueError:
        report_day = 4
    report_time_str = get_config("REPORT_TIME", "17:00")

    if weekday == report_day and is_time_due(report_time_str):
        already_sent = Sent.query.filter(
            Sent.target_date == today,
            Sent.kind.in_(["weekly_report", "weekly_report_warning"]),
            Sent.status == "success"
        ).first()

        if not already_sent:
            success, msg = dispatch_weekly_report(today)
            actions_executed.append({
                "action": "weekly_report",
                "target_date": today.isoformat(),
                "success": success,
                "message": msg
            })

    if not actions_executed:
        return jsonify({
            "status": "ok",
            "message": "Nothing due at this time",
            "current_time": now.isoformat(),
            "weekday": get_weekday_name(weekday)
        }), 200

    return jsonify({
        "status": "ok",
        "current_time": now.isoformat(),
        "weekday": get_weekday_name(weekday),
        "actions": actions_executed
    }), 200

# ==============================================================================
# Web Routes: Authentication
# ==============================================================================
@app.route("/login", methods=["GET", "POST"])
def login():
    """Single-user login protected by APP_PASSWORD with constant-time comparison."""
    if session.get("logged_in"):
        return redirect(url_for("dashboard"))

    error = None
    if request.method == "POST":
        entered_password = request.form.get("password", "").strip()
        app_password = get_config("APP_PASSWORD", "changeme").strip()
        if secrets.compare_digest(entered_password, app_password):
            session["logged_in"] = True
            session.permanent = True
            flash("Logged in successfully.", "success")
            next_url = request.args.get("next")
            return redirect(next_url or url_for("dashboard"))
        else:
            error = "Invalid password. Please check your APP_PASSWORD in .env."

    return render_template_string(HTML_TEMPLATE, view="login", error=error)


@app.route("/logout")
def logout():
    """Logs user out and clears session."""
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))

# ==============================================================================
# Web Routes: Dashboard & Main Features
# ==============================================================================
@app.route("/", methods=["GET"])
@login_required
def dashboard():
    """Main dashboard displaying notes, tomorrow's schedule, timetable manager, and settings."""
    now = get_now_user()
    today = now.date()
    tomorrow = today + timedelta(days=1)

    # 1. Daily Notes: Selected date handling
    selected_date_str = request.args.get("date", today.isoformat())
    try:
        selected_date = date.fromisoformat(selected_date_str)
    except Exception:
        selected_date = today

    monday, sunday, weekdays = get_week_range(today, include_weekend=True)
    notes_in_week = DailyNote.query.filter(DailyNote.date >= monday, DailyNote.date <= sunday).all()
    week_notes_dict = {n.date: n.content for n in notes_in_week}

    selected_note = DailyNote.query.filter_by(date=selected_date).first()
    selected_content = selected_note.content if selected_note else ""

    # Calendar cards: show recorded ONLY if non-empty content exists
    calendar_cards = []
    recorded_days_count = 0
    for d in weekdays:
        raw_note = week_notes_dict.get(d, "")
        has_note = bool(raw_note and raw_note.strip())
        if has_note:
            recorded_days_count += 1
        calendar_cards.append({
            "date": d,
            "iso": d.isoformat(),
            "name": d.strftime("%A"),
            "short_date": d.strftime("%d %b"),
            "is_today": (d == today),
            "is_selected": (d == selected_date),
            "has_note": has_note,
            "content": raw_note if has_note else ""
        })

    # 2. Tomorrow's Schedule & Human Input date helpers
    tomorrow_schedule = get_schedule_for_date(tomorrow)

    days_until_sat = (5 - today.weekday()) % 7
    this_saturday = today + timedelta(days=days_until_sat if days_until_sat > 0 else 7)
    days_until_sun = (6 - today.weekday()) % 7
    this_sunday = today + timedelta(days=days_until_sun if days_until_sun > 0 else 7)

    # Upcoming overrides (manual entries for upcoming 14 days, especially Saturday & Sunday)
    upcoming_overrides = ScheduleOverride.query.filter(
        ScheduleOverride.date >= today,
        ScheduleOverride.date <= today + timedelta(days=14)
    ).order_by(ScheduleOverride.date, ScheduleOverride.start_time).all()

    # 3. Semester Timetable (Recurring slots grouped by day of week, Mon=0 to Sun=6)
    all_slots = TimetableSlot.query.order_by(TimetableSlot.day_of_week, TimetableSlot.sort_order, TimetableSlot.start_time).all()
    slots_by_day = {i: [] for i in range(7)}  # 0=Mon..6=Sun
    for s in all_slots:
        if s.day_of_week in slots_by_day:
            slots_by_day[s.day_of_week].append(s)

    # 4. Latest timetable file/screenshot
    latest_img = TimetableImage.query.order_by(TimetableImage.uploaded_at.desc()).first()

    # 5. Live App Settings (Dynamically loaded from .env)
    settings_info = {
        "manager_email": get_config("MANAGER_EMAIL", "manager@example.com"),
        "owner_email": get_config("OWNER_EMAIL", "owner@example.com"),
        "timezone": get_config("TIMEZONE", "Asia/Kolkata"),
        "report_day": get_weekday_name(int(get_config("REPORT_DAY", "4"))),
        "report_time": get_config("REPORT_TIME", "17:00"),
        "schedule_time": get_config("SCHEDULE_TIME", "18:00"),
        "gemini_active": bool(get_config("GEMINI_API_KEY", "").strip()),
        "gemini_model": get_config("GEMINI_MODEL", "gemini-1.5-flash"),
        "timetable_excel_url": get_config("TIMETABLE_EXCEL_URL", ""),
        "smtp_configured": bool(get_config("SMTP_HOST") and get_config("SMTP_USER") and get_config("SMTP_PASS"))
    }

    return render_template_string(
        HTML_TEMPLATE,
        view="dashboard",
        csrf_token=get_csrf_token(),
        today=today,
        tomorrow=tomorrow,
        now=now,
        this_saturday=this_saturday,
        this_sunday=this_sunday,
        upcoming_overrides=upcoming_overrides,
        now_str=now.strftime("%A, %d %b %Y, %I:%M %p"),
        selected_date=selected_date,
        selected_content=selected_content,
        calendar_cards=calendar_cards,
        recorded_days_count=recorded_days_count,
        tomorrow_schedule=tomorrow_schedule,
        slots_by_day=slots_by_day,
        latest_img=latest_img,
        settings=settings_info
    )

# ==============================================================================
# Web Routes: Notes Actions
# ==============================================================================
@app.route("/notes/save", methods=["POST"])
@login_required
def save_note():
    """Saves daily note. Saving non-empty text creates/updates; empty text deletes."""
    validate_csrf()
    date_str = request.form.get("date", "").strip()
    content = request.form.get("content", "").strip()

    try:
        note_date = date.fromisoformat(date_str)
    except Exception:
        flash("Invalid date format.", "danger")
        return redirect(url_for("dashboard"))

    existing = DailyNote.query.filter_by(date=note_date).first()

    if not content:
        if existing:
            db.session.delete(existing)
            db.session.commit()
            flash(f"Note for {note_date.strftime('%A, %d %b %Y')} deleted.", "info")
        else:
            flash(f"No note existed for {note_date.strftime('%A, %d %b %Y')}.", "info")
    else:
        if existing:
            existing.content = content
            existing.updated_at = datetime.utcnow()
            flash(f"Note for {note_date.strftime('%A, %d %b %Y')} updated successfully.", "success")
        else:
            new_note = DailyNote(date=note_date, content=content)
            db.session.add(new_note)
            flash(f"Note for {note_date.strftime('%A, %d %b %Y')} saved successfully.", "success")
        db.session.commit()

    return redirect(url_for("dashboard", date=note_date.isoformat()))

# ==============================================================================
# Web Routes: Schedule Overrides (Human Input)
# ==============================================================================
@app.route("/schedule/override/add", methods=["POST"])
@login_required
def add_schedule_override():
    """Human input: adds an additional class or custom note for tomorrow or specific date."""
    validate_csrf()
    date_str = request.form.get("override_date", "").strip()
    action = request.form.get("action", "add")
    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()
    subject = request.form.get("subject", "").strip()
    room = request.form.get("room", "").strip()
    faculty = request.form.get("faculty", "").strip()
    notes = request.form.get("notes", "").strip()

    if not subject:
        flash("Please enter a subject or title for the class / note.", "danger")
        return redirect(url_for("dashboard"))

    try:
        target_d = date.fromisoformat(date_str)
    except Exception:
        target_d = get_now_user().date() + timedelta(days=1)

    override = ScheduleOverride(
        date=target_d,
        action=action,
        start_time=start_time,
        end_time=end_time,
        subject=subject,
        room=room,
        faculty=faculty,
        notes=notes
    )
    db.session.add(override)
    db.session.commit()

    flash(f"Added {subject} to schedule for {target_d.strftime('%A, %d %b %Y')}.", "success")
    return redirect(url_for("dashboard"))


@app.route("/schedule/override/delete/<int:override_id>", methods=["POST"])
@login_required
def delete_schedule_override(override_id):
    """Deletes an override entry."""
    validate_csrf()
    override = ScheduleOverride.query.get_or_404(override_id)
    db.session.delete(override)
    db.session.commit()
    flash("Additional schedule entry removed.", "info")
    return redirect(url_for("dashboard"))

# ==============================================================================
# Web Routes: Semester Timetable Slot Management
# ==============================================================================
@app.route("/timetable/slot/add", methods=["POST"])
@login_required
def add_timetable_slot():
    """Adds a recurring class slot to the semester timetable."""
    validate_csrf()
    try:
        day_of_week = int(request.form.get("day_of_week", 0))
    except ValueError:
        day_of_week = 0

    start_time = request.form.get("start_time", "").strip()
    end_time = request.form.get("end_time", "").strip()
    subject = request.form.get("subject", "").strip()
    room = request.form.get("room", "").strip()
    faculty = request.form.get("faculty", "").strip()

    if not subject or not start_time:
        flash("Subject and start time are required.", "danger")
        return redirect(url_for("dashboard"))

    new_slot = TimetableSlot(
        day_of_week=day_of_week,
        start_time=start_time,
        end_time=end_time or start_time,
        subject=subject,
        room=room,
        faculty=faculty
    )
    db.session.add(new_slot)
    db.session.commit()

    flash(f"Added recurring class '{subject}' to {get_weekday_name(day_of_week)}.", "success")
    return redirect(url_for("dashboard"))


@app.route("/timetable/slot/delete/<int:slot_id>", methods=["POST"])
@login_required
def delete_timetable_slot(slot_id):
    """Deletes a recurring class slot from the semester timetable."""
    validate_csrf()
    slot = TimetableSlot.query.get_or_404(slot_id)
    db.session.delete(slot)
    db.session.commit()
    flash(f"Removed '{slot.subject}' from {get_weekday_name(slot.day_of_week)} schedule.", "info")
    return redirect(url_for("dashboard"))

# ==============================================================================
# Web Routes: Timetable Excel Import, Image Upload & REST APIs
# ==============================================================================
@app.route("/timetable/upload", methods=["POST"])
@login_required
def upload_timetable():
    """
    Uploads timetable file (Excel .xlsx, .csv, or image .png/.jpg/.webp).
    - If Excel/CSV: parses directly via openpyxl / csv into TimetableSlot.
    - If image: attempts Gemini AI Vision auto-extraction.
    - Saves file in database for download or preview.
    """
    validate_csrf()
    file = request.files.get("timetable_file") or request.files.get("timetable_image")

    if not file or not file.filename:
        flash("No file selected.", "danger")
        return redirect(url_for("dashboard"))

    allowed_exts = {".xlsx", ".xls", ".csv", ".png", ".jpg", ".jpeg", ".webp"}
    filename = file.filename
    _, ext = os.path.splitext(filename.lower())
    if ext not in allowed_exts:
        flash("Supported formats: Excel (.xlsx, .xls), CSV (.csv), or Image (.png, .jpg, .webp).", "danger")
        return redirect(url_for("dashboard"))

    mimetype = file.mimetype or "application/octet-stream"
    file_bytes = file.read()
    if len(file_bytes) > 10 * 1024 * 1024:
        flash("File exceeds maximum allowed size of 10 MB.", "danger")
        return redirect(url_for("dashboard"))

    # Store latest file in database
    TimetableImage.query.delete()
    new_file = TimetableImage(
        filename=filename,
        mimetype=mimetype,
        data=file_bytes,
        uploaded_at=datetime.utcnow()
    )
    db.session.add(new_file)
    db.session.commit()

    if ext in {".xlsx", ".xls", ".csv"}:
        count, msg = parse_timetable_excel(file_bytes, filename=filename)
        if count and count > 0:
            flash(f"Excel timetable imported! {msg}", "success")
        else:
            flash(f"File saved. ({msg})", "info")
    else:
        count, msg = extract_timetable_from_image(file_bytes, mimetype)
        if count and count > 0:
            flash(f"Timetable uploaded! {msg}", "success")
        else:
            flash(f"Timetable file saved. ({msg})", "info")

    return redirect(url_for("dashboard"))


@app.route("/timetable/download-excel")
@login_required
def download_timetable_excel():
    """Generates and downloads current semester timetable as an Excel file."""
    slots = TimetableSlot.query.order_by(TimetableSlot.day_of_week, TimetableSlot.sort_order, TimetableSlot.start_time).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Timetable"

    header_fill = openpyxl.styles.PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = openpyxl.styles.Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    day_fill = openpyxl.styles.PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    day_font = openpyxl.styles.Font(name="Calibri", size=11, bold=True, color="0F172A")
    thin_border = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style="thin", color="CBD5E1"),
        right=openpyxl.styles.Side(style="thin", color="CBD5E1"),
        top=openpyxl.styles.Side(style="thin", color="CBD5E1"),
        bottom=openpyxl.styles.Side(style="thin", color="CBD5E1")
    )

    headers = ["Day", "8:00- 9:00", "9:00-10:00", "10:00-11:00", "11:00-12:00", "12:00-1:00", "1:00-2:00", "2:00-3:00", "3:00-4:00"]
    ws.append(headers)

    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    for d_idx, day_name in enumerate(days):
        day_slots = [s for s in slots if s.day_of_week == d_idx]
        row_vals = [day_name]
        for h in headers[1:]:
            st_norm, et_norm = normalize_time_slot(h)
            matching = [s for s in day_slots if s.start_time.strip() == st_norm]
            if matching:
                val = f"{matching[0].subject}" + (f", {matching[0].room}" if matching[0].room else "")
                row_vals.append(val)
            else:
                row_vals.append("")
        ws.append(row_vals)

    ws.append([])
    ws.append(["Acronym", "Full Course Name"])
    ws.append(["DMW", "Data Mining and Warehousing"])
    ws.append(["GEN AI", "Generative Models and Application"])
    ws.append(["DLMA", "Deep Learning Models and Application"])
    ws.append(["NLP", "Natural Language Processing"])

    # Style header
    for col_num in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_num)
        c.fill = header_fill
        c.font = header_font
        c.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")

    for row_num in range(2, 9):
        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=row_num, column=col_num)
            c.border = thin_border
            if col_num == 1:
                c.fill = day_fill
                c.font = day_font
            else:
                c.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
                if c.value:
                    c.fill = openpyxl.styles.PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
                    c.font = openpyxl.styles.Font(name="Calibri", size=10, bold=True, color="1E3A8A")

    ws.column_dimensions["A"].width = 14
    for col_letter in ["B", "C", "D", "E", "F", "G", "H", "I"]:
        ws.column_dimensions[col_letter].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="semester_timetable.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/timetable/preview")
@login_required
def preview_timetable():
    """Serves the latest timetable screenshot or file."""
    img = TimetableImage.query.order_by(TimetableImage.uploaded_at.desc()).first()
    if not img or not img.data:
        abort(404, description="No timetable file has been uploaded.")
    return Response(img.data, mimetype=img.mimetype)


# ==============================================================================
# REST API Endpoints: Timetable Upload & Slots
# ==============================================================================
@app.route("/api/timetable/upload", methods=["POST"])
def api_upload_timetable():
    """
    REST API endpoint to upload and parse an Excel or image timetable.
    Accepts multipart/form-data with 'file'.
    """
    file = request.files.get("file")
    if not file:
        return jsonify({"status": "error", "message": "Missing 'file' field in multipart form-data"}), 400

    filename = file.filename
    _, ext = os.path.splitext(filename.lower())
    file_bytes = file.read()
    mimetype = file.mimetype or "application/octet-stream"

    # Save to database
    TimetableImage.query.delete()
    new_file = TimetableImage(
        filename=filename,
        mimetype=mimetype,
        data=file_bytes,
        uploaded_at=datetime.utcnow()
    )
    db.session.add(new_file)
    db.session.commit()

    if ext in {".xlsx", ".xls", ".csv"}:
        count, msg = parse_timetable_excel(file_bytes, filename=filename)
    else:
        count, msg = extract_timetable_from_image(file_bytes, mimetype)

    if count:
        slots = TimetableSlot.query.order_by(TimetableSlot.day_of_week, TimetableSlot.sort_order, TimetableSlot.start_time).all()
        slots_json = [{
            "id": s.id,
            "day_of_week": s.day_of_week,
            "day_name": get_weekday_name(s.day_of_week),
            "start_time": s.start_time,
            "end_time": s.end_time,
            "subject": s.subject,
            "room": s.room,
            "faculty": s.faculty
        } for s in slots]
        return jsonify({"status": "success", "count": count, "message": msg, "slots": slots_json}), 200
    else:
        return jsonify({"status": "error", "message": msg}), 400


@app.route("/api/timetable/slots", methods=["GET", "POST"])
def api_timetable_slots():
    """
    REST API endpoint to retrieve or create semester timetable slots.
    """
    if request.method == "POST":
        data = request.get_json(silent=True) or request.form
        try:
            day_of_week = int(data.get("day_of_week", 0))
        except (ValueError, TypeError):
            day_of_week = 0

        subject = data.get("subject", "").strip()
        start_time = data.get("start_time", "").strip()
        end_time = data.get("end_time", "").strip() or start_time
        room = data.get("room", "").strip()
        faculty = data.get("faculty", "").strip()

        if not subject or not start_time:
            return jsonify({"status": "error", "message": "subject and start_time are required"}), 400

        new_slot = TimetableSlot(
            day_of_week=day_of_week,
            start_time=start_time,
            end_time=end_time,
            subject=subject,
            room=room,
            faculty=faculty
        )
        db.session.add(new_slot)
        db.session.commit()
        return jsonify({"status": "success", "id": new_slot.id, "day": get_weekday_name(day_of_week)}), 201

    slots = TimetableSlot.query.order_by(TimetableSlot.day_of_week, TimetableSlot.sort_order, TimetableSlot.start_time).all()
    slots_json = [{
        "id": s.id,
        "day_of_week": s.day_of_week,
        "day_name": get_weekday_name(s.day_of_week),
        "start_time": s.start_time,
        "end_time": s.end_time,
        "subject": s.subject,
        "room": s.room,
        "faculty": s.faculty
    } for s in slots]
    return jsonify({"status": "success", "count": len(slots), "slots": slots_json}), 200


# ==============================================================================
# Web Routes: Instant Dispatch & Previews
# ==============================================================================
@app.route("/schedule/send-now", methods=["POST"])
@login_required
def send_schedule_now():
    """Manual button: immediately sends tomorrow's (or chosen date's) schedule to manager (CC owner)."""
    validate_csrf()
    recipient = request.form.get("recipient_email", "").strip()
    if recipient:
        if recipient != get_config("MANAGER_EMAIL"):
            update_env_file({"MANAGER_EMAIL": recipient})
    else:
        recipient = get_config("MANAGER_EMAIL")

    target_date_str = request.form.get("target_date", "").strip()
    try:
        target_date = date.fromisoformat(target_date_str) if target_date_str else (get_now_user().date() + timedelta(days=1))
    except Exception:
        target_date = get_now_user().date() + timedelta(days=1)

    success, msg = dispatch_daily_schedule(target_date, recipient_email=recipient)
    flash(msg, "success" if success else "danger")
    return redirect(url_for("dashboard"))


@app.route("/schedule/preview-email")
@login_required
def preview_schedule_email():
    """Previews the rendered HTML email for tomorrow's schedule in browser."""
    date_str = request.args.get("date")
    try:
        target_date = date.fromisoformat(date_str) if date_str else (get_now_user().date() + timedelta(days=1))
    except Exception:
        target_date = get_now_user().date() + timedelta(days=1)

    schedule_data = get_schedule_for_date(target_date)
    return render_template_string(build_schedule_email_html(schedule_data))


@app.route("/report/send-now", methods=["POST"])
@login_required
def send_report_now():
    """Manual button: sends weekly Excel report now."""
    validate_csrf()
    today = get_now_user().date()
    success, msg = dispatch_weekly_report(today)
    flash(msg, "success" if success else "danger")
    return redirect(url_for("dashboard"))


@app.route("/report/download")
@login_required
def download_preview_report():
    """Generates preview Excel report for current week and downloads it directly."""
    ref_date_str = request.args.get("date")
    ref_date = date.fromisoformat(ref_date_str) if ref_date_str else get_now_user().date()

    data = build_week_report_data(ref_date)
    excel_buf = generate_weekly_report_excel(
        data["monday"],
        data["friday"],
        data["parsed_tasks"],
        data["weekdays_meta"]
    )
    filename = f"Weekly_Report_Preview_{data['monday'].isoformat()}_to_{data['friday'].isoformat()}.xlsx"

    return send_file(
        excel_buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/timetable/sync-url", methods=["POST"])
@login_required
def sync_timetable_url():
    """Manual trigger: syncs timetable from configured or submitted cloud URL."""
    validate_csrf()
    cloud_url = request.form.get("timetable_excel_url", "").strip()
    if cloud_url:
        update_env_file({"TIMETABLE_EXCEL_URL": cloud_url})
    else:
        cloud_url = get_config("TIMETABLE_EXCEL_URL", "").strip()

    if not cloud_url:
        flash("Please provide a Google Sheets or OneDrive link first.", "warning")
        return redirect(url_for("dashboard"))

    count, msg = sync_timetable_from_url(cloud_url)
    if count and count > 0:
        flash(f"{msg}", "success")
    else:
        flash(f"{msg}", "danger")
    return redirect(url_for("dashboard"))


@app.route("/settings/save", methods=["POST"])
@login_required
def save_settings():
    """Allows user to update manager email, send times, and other settings directly in UI."""
    validate_csrf()
    new_settings = {
        "MANAGER_EMAIL": request.form.get("manager_email", "").strip(),
        "OWNER_EMAIL": request.form.get("owner_email", "").strip(),
        "SCHEDULE_TIME": request.form.get("schedule_time", "18:00").strip(),
        "REPORT_TIME": request.form.get("report_time", "17:00").strip(),
        "REPORT_DAY": request.form.get("report_day", "4").strip(),
        "TIMEZONE": request.form.get("timezone", "Asia/Kolkata").strip(),
        "TIMETABLE_EXCEL_URL": request.form.get("timetable_excel_url", "").strip()
    }

    gemini_key = request.form.get("gemini_api_key", "").strip()
    if gemini_key:
        new_settings["GEMINI_API_KEY"] = gemini_key

    update_env_file(new_settings)
    flash("Settings updated successfully! Changes are live immediately.", "success")
    return redirect(url_for("dashboard"))

# ==============================================================================
# UI HTML/CSS/JS Template
# ==============================================================================
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Work Reporter</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-body: #f8fafc;
            --bg-card: #ffffff;
            --bg-card-subtle: #f1f5f9;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --border: #e2e8f0;
            --primary: #2563eb;
            --primary-hover: #1d4ed8;
            --primary-light: #eff6ff;
            --success-bg: #ecfdf5;
            --success-text: #065f46;
            --success-border: #a7f3d0;
            --warning-bg: #fffbeb;
            --warning-text: #92400e;
            --warning-border: #fde68a;
            --danger-bg: #fef2f2;
            --danger-text: #991b1b;
            --danger-border: #fecaca;
            --info-bg: #f0f9ff;
            --info-text: #075985;
            --info-border: #bae6fd;
            --shadow-sm: 0 1px 2px 0 rgb(0 0 0 / 0.05);
            --shadow: 0 4px 6px -1px rgb(0 0 0 / 0.08), 0 2px 4px -2px rgb(0 0 0 / 0.05);
            --radius: 10px;
        }

        [data-theme="dark"] {
            --bg-body: #0b1120;
            --bg-card: #1e293b;
            --bg-card-subtle: #0f172a;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --border: #334155;
            --primary: #3b82f6;
            --primary-hover: #60a5fa;
            --primary-light: #1e3a8a33;
            --success-bg: #064e3b33;
            --success-text: #6ee7b7;
            --success-border: #065f46;
            --warning-bg: #78350f33;
            --warning-text: #fcd34d;
            --warning-border: #92400e;
            --danger-bg: #7f1d1d33;
            --danger-text: #fca5a5;
            --danger-border: #991b1b;
            --info-bg: #0c4a6e33;
            --info-text: #7dd3fc;
            --info-border: #0369a1;
            --shadow: 0 4px 6px -1px rgb(0 0 0 / 0.3);
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            background-color: var(--bg-body);
            color: var(--text-main);
            line-height: 1.5;
            transition: background-color 0.2s, color 0.2s;
            min-height: 100vh;
            display: flex;
            flex-direction: column;
        }

        .navbar {
            background-color: var(--bg-card);
            border-bottom: 1px solid var(--border);
            padding: 12px 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: var(--shadow-sm);
        }

        .navbar-brand {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 1.25rem;
            font-weight: 700;
            color: var(--primary);
            text-decoration: none;
        }

        .navbar-brand svg { width: 26px; height: 26px; fill: currentColor; }

        .navbar-actions { display: flex; align-items: center; gap: 14px; }

        .theme-toggle, .btn-logout {
            background: var(--bg-card-subtle);
            border: 1px solid var(--border);
            color: var(--text-main);
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 0.85rem;
            cursor: pointer;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-weight: 500;
            transition: all 0.15s;
        }

        .theme-toggle:hover, .btn-logout:hover {
            border-color: var(--primary);
            color: var(--primary);
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 24px 20px;
            width: 100%;
            flex: 1;
        }

        .flashes { list-style: none; margin-bottom: 20px; }
        .alert {
            padding: 12px 16px;
            border-radius: var(--radius);
            margin-bottom: 10px;
            font-size: 0.95rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
            border: 1px solid transparent;
        }
        .alert-success { background: var(--success-bg); color: var(--success-text); border-color: var(--success-border); }
        .alert-warning { background: var(--warning-bg); color: var(--warning-text); border-color: var(--warning-border); }
        .alert-danger { background: var(--danger-bg); color: var(--danger-text); border-color: var(--danger-border); }
        .alert-info { background: var(--info-bg); color: var(--info-text); border-color: var(--info-border); }

        .login-wrap {
            max-width: 400px;
            margin: 80px auto;
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 32px;
            box-shadow: var(--shadow);
        }
        .login-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 8px; color: var(--primary); text-align: center; }
        .login-subtitle { font-size: 0.9rem; color: var(--text-muted); margin-bottom: 24px; text-align: center; }

        .dashboard-grid {
            display: grid;
            grid-template-columns: 1.8fr 1.2fr;
            gap: 24px;
        }
        @media (max-width: 950px) {
            .dashboard-grid { grid-template-columns: 1fr; }
        }

        .card {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 20px;
            margin-bottom: 24px;
            box-shadow: var(--shadow-sm);
        }
        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            border-bottom: 1px solid var(--border);
            padding-bottom: 12px;
        }
        .card-title {
            font-size: 1.05rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .badge {
            display: inline-flex;
            align-items: center;
            padding: 3px 9px;
            border-radius: 9999px;
            font-size: 0.75rem;
            font-weight: 600;
        }
        .badge-success { background: var(--success-bg); color: var(--success-text); border: 1px solid var(--success-border); }
        .badge-warning { background: var(--warning-bg); color: var(--warning-text); border: 1px solid var(--warning-border); }
        .badge-info { background: var(--info-bg); color: var(--info-text); border: 1px solid var(--info-border); }

        .week-cards {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 10px;
            margin-bottom: 20px;
        }
        @media (max-width: 960px) {
            .week-cards { grid-template-columns: repeat(4, 1fr); }
        }
        @media (max-width: 600px) {
            .week-cards { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 380px) {
            .week-cards { grid-template-columns: 1fr; }
        }
        .day-card {
            background: var(--bg-card-subtle);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 12px;
            text-align: center;
            cursor: pointer;
            transition: all 0.15s;
            text-decoration: none;
            color: inherit;
            display: flex;
            flex-direction: column;
            gap: 4px;
        }
        .day-card:hover { border-color: var(--primary); transform: translateY(-1px); }
        .day-card.selected { border-color: var(--primary); background: var(--primary-light); outline: 2px solid var(--primary); }
        .day-card.today {
            border-color: var(--primary);
            box-shadow: 0 0 0 2px var(--primary);
            position: relative;
        }
        .day-card.today .day-card-name {
            color: var(--primary);
            font-weight: 700;
        }
        .day-card-name { font-size: 0.85rem; font-weight: 600; color: var(--text-muted); }
        .day-card.selected .day-card-name { color: var(--primary); }
        .day-card-date { font-size: 1rem; font-weight: 700; }

        .form-group { margin-bottom: 14px; }
        .form-row { display: flex; gap: 10px; margin-bottom: 14px; }
        .form-row .form-group { margin-bottom: 0; flex: 1; }
        .form-label { display: block; font-size: 0.85rem; font-weight: 600; margin-bottom: 5px; color: var(--text-main); }
        .form-control {
            width: 100%;
            padding: 9px 12px;
            font-size: 0.9rem;
            border: 1px solid var(--border);
            border-radius: 6px;
            background: var(--bg-card);
            color: var(--text-main);
            outline: none;
            font-family: inherit;
            transition: border-color 0.15s;
        }
        .form-control:focus { border-color: var(--primary); }
        textarea.form-control { resize: vertical; min-height: 110px; }

        .btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            padding: 8px 16px;
            border-radius: 6px;
            font-size: 0.875rem;
            font-weight: 600;
            border: none;
            cursor: pointer;
            text-decoration: none;
            transition: all 0.15s;
        }
        .btn-primary { background: var(--primary); color: #ffffff; }
        .btn-primary:hover { background: var(--primary-hover); }
        .btn-outline { background: transparent; border: 1px solid var(--border); color: var(--text-main); }
        .btn-outline:hover { border-color: var(--primary); color: var(--primary); }
        .btn-sm { padding: 4px 8px; font-size: 0.75rem; }
        .btn-block { width: 100%; }
        .btn-group { display: flex; gap: 10px; margin-top: 10px; }

        .note-preview {
            background: var(--bg-card-subtle);
            border-left: 4px solid var(--primary);
            padding: 12px 14px;
            border-radius: 0 6px 6px 0;
            font-size: 0.9rem;
            margin-top: 12px;
            white-space: pre-wrap;
            word-break: break-word;
        }

        .schedule-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 12px;
            border-radius: 6px;
            background: var(--bg-card-subtle);
            margin-bottom: 8px;
            border: 1px solid var(--border);
        }
        .schedule-item.additional {
            background: var(--warning-bg);
            border-color: var(--warning-border);
        }
        .schedule-time {
            font-weight: 700;
            color: var(--primary);
            font-size: 0.85rem;
            min-width: 140px;
        }
        .schedule-item.additional .schedule-time { color: var(--warning-text); }
        .schedule-detail { flex: 1; }
        .schedule-subject { font-weight: 600; font-size: 0.9rem; }
        .schedule-meta { font-size: 0.75rem; color: var(--text-muted); }

        .tab-nav {
            display: flex;
            gap: 6px;
            border-bottom: 1px solid var(--border);
            margin-bottom: 14px;
            overflow-x: auto;
        }
        .tab-btn {
            background: none;
            border: none;
            border-bottom: 2px solid transparent;
            padding: 8px 12px;
            font-size: 0.85rem;
            font-weight: 600;
            color: var(--text-muted);
            cursor: pointer;
            white-space: nowrap;
        }
        .tab-btn.active {
            color: var(--primary);
            border-bottom-color: var(--primary);
        }

        .info-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
        .info-table td { padding: 7px 0; border-bottom: 1px solid var(--border); }
        .info-table tr:last-child td { border-bottom: none; }
        .info-label { color: var(--text-muted); width: 45%; }
        .info-value { font-weight: 500; text-align: right; word-break: break-all; }

        .footer {
            text-align: center;
            font-size: 0.8rem;
            color: var(--text-muted);
            padding: 20px 0;
            border-top: 1px solid var(--border);
            background: var(--bg-card);
        }
    </style>
</head>
<body>

    <!-- Header -->
    <header class="navbar">
        <a href="{{ url_for('dashboard') }}" class="navbar-brand">
            <svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>
            Work Reporter
        </a>
        <div class="navbar-actions">
            {% if session.get('logged_in') and now %}
                <span class="badge badge-info" style="font-size: 0.82rem; padding: 6px 12px; display: inline-flex; align-items: center; gap: 6px; font-weight: 600;">
                    🗓️ <strong>{{ now.strftime('%A, %d %b %Y') }}</strong>
                </span>
            {% endif %}
            <button type="button" class="theme-toggle" id="themeToggle" title="Toggle Theme">
                <span id="themeIcon">🌓</span>
                <span id="themeLabel">Theme</span>
            </button>
            {% if session.get('logged_in') %}
                <a href="{{ url_for('logout') }}" class="btn-logout">Logout</a>
            {% endif %}
        </div>
    </header>

    <main class="container">

        <!-- Flashes -->
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                <ul class="flashes">
                {% for category, message in messages %}
                    <li class="alert alert-{{ category }}">
                        <span>{{ message }}</span>
                        <button type="button" onclick="this.parentElement.remove()" style="background:none; border:none; color:inherit; cursor:pointer;">✕</button>
                    </li>
                {% endfor %}
                </ul>
            {% endif %}
        {% endwith %}

        {% if view == 'login' %}
            <!-- Login Card -->
            <div class="login-wrap">
                <h1 class="login-title">Work Reporter</h1>
                <p class="login-subtitle">Sign in to your work reporter dashboard</p>

                {% if error %}
                    <div class="alert alert-danger" style="margin-bottom: 16px;">
                        {{ error }}
                    </div>
                {% endif %}

                <form method="POST" action="{{ url_for('login') }}">
                    <div class="form-group">
                        <label class="form-label" for="password">Password</label>
                        <input type="password" id="password" name="password" class="form-control" placeholder="Enter APP_PASSWORD" required autofocus>
                    </div>
                    <button type="submit" class="btn btn-primary btn-block">Sign In</button>
                </form>
            </div>

        {% elif view == 'dashboard' %}

            <!-- Top Row: Weekly Notes (Mon-Sun) -->
            <div class="card" style="margin-bottom: 20px;">
                <div class="card-header">
                    <div>
                        <span class="card-title">📅 Weekly Notes (Mon – Sun)</span>
                        <span style="font-size: 0.8rem; color: var(--text-muted); margin-left: 8px;">Timezone: <strong>{{ settings.timezone }}</strong></span>
                    </div>
                    <span class="badge badge-info">{{ recorded_days_count }} / {{ calendar_cards|length }} Days Recorded</span>
                </div>

                <div class="week-cards">
                    {% for c in calendar_cards %}
                        <a href="{{ url_for('dashboard', date=c.iso) }}" class="day-card {% if c.is_selected %}selected{% endif %} {% if c.is_today %}today{% endif %}" onclick="selectDateCard('{{ c.iso }}', {{ c.content|tojson }})">
                            <span class="day-card-name">
                                {{ c.name }}
                                {% if c.is_today %}
                                    <span style="background: var(--primary); color: white; font-size: 0.65rem; padding: 2px 6px; border-radius: 4px; margin-left: 4px; vertical-align: middle;">Today</span>
                                {% endif %}
                            </span>
                            <span class="day-card-date">{{ c.short_date }}</span>
                            <span class="day-card-status">
                                {% if c.has_note %}
                                    <span class="badge badge-success">✓ Recorded</span>
                                {% else %}
                                    <span class="badge badge-warning">Empty</span>
                                {% endif %}
                            </span>
                        </a>
                    {% endfor %}
                </div>
            </div>

            <!-- Main Layout Grid -->
            <div class="dashboard-grid">

                <!-- Left Column: Daily Notes + Tomorrow's Schedule (Human Input) -->
                <div>
                    <!-- Daily Notes Editor Card -->
                    <div class="card">
                        <div class="card-header">
                            <span class="card-title">📝 Daily Note Entry</span>
                            <span class="badge badge-info">{{ selected_date.strftime('%A, %d %b %Y') }}</span>
                        </div>

                        <form method="POST" action="{{ url_for('save_note') }}">
                            <input type="hidden" name="csrf_token" value="{{ csrf_token }}">

                            <div class="form-group">
                                <label class="form-label" for="datePicker">Date</label>
                                <input type="date" id="datePicker" name="date" class="form-control" value="{{ selected_date.isoformat() }}" onchange="handleDateChange(this.value)" required>
                            </div>

                            <div class="form-group">
                                <label class="form-label" for="noteContent">Daily Work Notes (Plain English)</label>
                                <textarea id="noteContent" name="content" class="form-control" rows="5" placeholder="E.g. Fixed login bug, attended sprint planning for 1 hr, reviewed 2 PRs">{{ selected_content }}</textarea>
                                <p style="font-size: 0.8rem; color: var(--text-muted); margin-top: 6px;">
                                    💡 <strong>Rule:</strong> Entering notes marks the day as Recorded. Leaving text blank and saving deletes the note, returning the day to Empty.
                                </p>
                            </div>

                            <div class="btn-group">
                                <button type="submit" class="btn btn-primary">Save Note</button>
                                <button type="button" class="btn btn-outline" onclick="document.getElementById('noteContent').value = '';">Clear</button>
                            </div>
                        </form>

                        {% if selected_content %}
                            <div style="margin-top: 16px;">
                                <span class="form-label" style="font-size: 0.8rem; color: var(--text-muted);">Current Saved Note:</span>
                                <div class="note-preview">{{ selected_content }}</div>
                            </div>
                        {% endif %}
                    </div>

                    <!-- Tomorrow's Schedule & Human Input Area -->
                    <div class="card" style="border: 1.5px solid var(--primary);">
                        <div class="card-header" style="background: var(--primary-light); margin: -20px -20px 16px -20px; padding: 16px 20px; border-radius: var(--radius) var(--radius) 0 0;">
                            <div>
                                <span class="card-title" style="color: var(--primary);">
                                    ⏰ Tomorrow's Schedule: {{ tomorrow_schedule.weekday_name }}, {{ tomorrow.strftime('%d %b %Y') }}
                                </span>
                                <p style="font-size: 0.8rem; color: var(--text-muted); margin-top: 2px;">
                                    Dispatched every night at <strong>{{ settings.schedule_time }}</strong> to <strong>{{ settings.manager_email }}</strong>
                                </p>
                            </div>
                            <span class="badge badge-success">Ready to Send</span>
                        </div>

                        <!-- Schedule Items for Tomorrow -->
                        <div style="margin-bottom: 16px;">
                            <h4 style="font-size: 0.9rem; margin-bottom: 10px; color: var(--text-main);">Scheduled Classes & Events:</h4>

                            {% if tomorrow_schedule.regular_classes or tomorrow_schedule.additional_classes %}
                                {% for r in tomorrow_schedule.regular_classes %}
                                    <div class="schedule-item">
                                        <div class="schedule-time">🕒 {{ r.start_time }} – {{ r.end_time }}</div>
                                        <div class="schedule-detail">
                                            <div class="schedule-subject">{{ r.subject }}</div>
                                            <div class="schedule-meta">Room: {{ r.room or '—' }} | Faculty: {{ r.faculty or '—' }}</div>
                                        </div>
                                    </div>
                                {% endfor %}

                                {% for a in tomorrow_schedule.additional_classes %}
                                    <div class="schedule-item additional">
                                        <div class="schedule-time">⭐ {{ a.start_time }} – {{ a.end_time }}</div>
                                        <div class="schedule-detail">
                                            <div class="schedule-subject">{{ a.subject }} <span class="badge badge-warning" style="font-size: 0.65rem;">EXTRA CLASS</span></div>
                                            <div class="schedule-meta">Room: {{ a.room or '—' }} | Faculty: {{ a.faculty or '—' }}{% if a.notes %} | Note: {{ a.notes }}{% endif %}</div>
                                        </div>
                                        <form method="POST" action="{{ url_for('delete_schedule_override', override_id=a.id) }}" style="display:inline;">
                                            <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
                                            <button type="submit" class="btn btn-outline btn-sm" title="Remove additional class" onclick="return confirm('Remove this extra class?');">✕</button>
                                        </form>
                                    </div>
                                {% endfor %}
                            {% else %}
                                <div style="background: var(--bg-card-subtle); padding: 14px; border-radius: 6px; text-align: center; color: var(--text-muted); font-size: 0.9rem;">
                                    No regular classes scheduled for {{ tomorrow_schedule.weekday_name }}. You can add an extra class below!
                                </div>
                            {% endif %}
                        </div>

                        <!-- Human Input: Add Class / Special Note (Tomorrow, Saturday, Sunday, or any date) -->
                        <div style="background: var(--bg-card-subtle); padding: 14px; border-radius: 8px; margin-bottom: 16px;">
                            <strong style="display: block; font-size: 0.85rem; margin-bottom: 8px; color: var(--text-main);">
                                ➕ Human Input: Add Class or Special Note (Any Date / Saturday / Sunday)
                            </strong>
                            <form method="POST" action="{{ url_for('add_schedule_override') }}">
                                <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
                                <input type="hidden" name="action" value="add">

                                <div class="form-group" style="margin-bottom: 10px;">
                                    <label class="form-label" style="font-weight: 600;">Target Date:</label>
                                    <div style="display: flex; gap: 8px; align-items: center; flex-wrap: wrap;">
                                        <input type="date" name="override_date" id="overrideDateInput" class="form-control" value="{{ tomorrow.isoformat() }}" style="max-width: 170px;" required>
                                        <button type="button" class="btn btn-outline btn-sm" onclick="setOverrideDate('{{ tomorrow.isoformat() }}')">Tomorrow ({{ tomorrow.strftime('%a') }})</button>
                                        <button type="button" class="btn btn-outline btn-sm" onclick="setOverrideDate('{{ this_saturday.isoformat() }}')">This Sat ({{ this_saturday.strftime('%d %b') }})</button>
                                        <button type="button" class="btn btn-outline btn-sm" onclick="setOverrideDate('{{ this_sunday.isoformat() }}')">This Sun ({{ this_sunday.strftime('%d %b') }})</button>
                                    </div>
                                </div>

                                <div class="form-row">
                                    <div class="form-group">
                                        <label class="form-label">Subject / Topic</label>
                                        <input type="text" name="subject" class="form-control" placeholder="E.g. Data Mining / Remedial Session" required>
                                    </div>
                                    <div class="form-group">
                                        <label class="form-label">Section / Venue</label>
                                        <input type="text" name="room" class="form-control" placeholder="E.g. MCA B / Room 204">
                                    </div>
                                </div>

                                <div class="form-row">
                                    <div class="form-group">
                                        <label class="form-label">Start Time</label>
                                        <input type="text" name="start_time" class="form-control" placeholder="E.g. 10:00 AM">
                                    </div>
                                    <div class="form-group">
                                        <label class="form-label">End Time</label>
                                        <input type="text" name="end_time" class="form-control" placeholder="E.g. 11:30 AM">
                                    </div>
                                    <div class="form-group">
                                        <label class="form-label">Instructor / Note</label>
                                        <input type="text" name="notes" class="form-control" placeholder="E.g. Special weekend session">
                                    </div>
                                </div>

                                <button type="submit" class="btn btn-outline btn-sm">Add to Schedule</button>
                            </form>

                            {% if upcoming_overrides %}
                                <div style="margin-top: 14px; border-top: 1px dashed var(--border); padding-top: 10px;">
                                    <div style="font-size: 0.8rem; font-weight: 700; color: var(--text-main); margin-bottom: 8px;">
                                        📅 Upcoming Scheduled Manual / Weekend Classes ({{ upcoming_overrides|length }})
                                    </div>
                                    {% for ov in upcoming_overrides %}
                                        <div class="schedule-item additional" style="margin-bottom: 6px;">
                                            <div class="schedule-time">
                                                <div style="font-weight: 700;">{{ ov.date.strftime('%a, %d %b') }}</div>
                                                <small style="color: #b45309;">{{ ov.start_time or 'TBD' }} – {{ ov.end_time or 'TBD' }}</small>
                                            </div>
                                            <div class="schedule-detail">
                                                <div class="schedule-subject">{{ ov.subject }} <span class="badge badge-warning" style="font-size: 0.65rem;">MANUAL ENTRY</span></div>
                                                <div class="schedule-meta">Room/Sec: {{ ov.room or '—' }} {% if ov.notes %}| Note: {{ ov.notes }}{% endif %}</div>
                                            </div>
                                            <form method="POST" action="{{ url_for('delete_schedule_override', override_id=ov.id) }}" style="display:inline;">
                                                <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
                                                <button type="submit" class="btn btn-outline btn-sm" title="Remove" onclick="return confirm('Remove this class?');">✕</button>
                                            </form>
                                        </div>
                                    {% endfor %}
                                </div>
                            {% endif %}
                        </div>

                        <!-- Send Schedule Form with Editable Recipient -->
                        <form method="POST" action="{{ url_for('send_schedule_now') }}" style="background: var(--bg-card-subtle); padding: 14px; border-radius: 8px; border: 1px solid var(--border); margin-top: 10px;">
                            <input type="hidden" name="csrf_token" value="{{ csrf_token }}">

                            <div class="form-group" style="margin-bottom: 10px;">
                                <label class="form-label" style="font-weight: 700; color: var(--primary);">
                                    ✉️ Send Daily Schedule To (Recipient Email):
                                </label>
                                <div style="display: flex; gap: 8px; flex-wrap: wrap;">
                                    <input type="email" name="recipient_email" class="form-control" value="{{ settings.manager_email }}" style="flex: 1; min-width: 220px; font-weight: 600;" required>
                                    <button type="submit" class="btn btn-primary" style="white-space: nowrap;" onclick="return confirm('Send schedule now?');">
                                        🚀 Send Schedule Now
                                    </button>
                                </div>
                                <p style="font-size: 0.75rem; color: var(--text-muted); margin-top: 4px;">
                                    💡 Prefilled with your manager email. You can change this address anytime right here.
                                </p>
                            </div>

                            <div style="display: flex; gap: 10px; align-items: center; margin-top: 8px; flex-wrap: wrap;">
                                <span style="font-size: 0.8rem; color: var(--text-muted); font-weight: 600;">Target Date:</span>
                                <input type="date" name="target_date" value="{{ tomorrow.isoformat() }}" class="form-control" style="max-width: 160px; height: 32px; font-size: 0.8rem;">
                                <a href="{{ url_for('preview_schedule_email', date=tomorrow.isoformat()) }}" target="_blank" class="btn btn-outline btn-sm">
                                    👁️ Preview Email (What Manager Receives)
                                </a>
                                <span style="font-size: 0.75rem; color: var(--text-muted);">
                                    • Sends simple message on: Date, Class, and Timings
                                </span>
                            </div>
                        </form>
                    </div>

                    <!-- Weekly Excel Actions Card -->
                    <div class="card">
                        <div class="card-header">
                            <span class="card-title">📊 Weekly Excel Report</span>
                            <span class="badge badge-info">Sent {{ settings.report_day }} @ {{ settings.report_time }}</span>
                        </div>
                        <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 12px;">
                            Parses daily notes into categorized tasks, calculates hours, and prepares the <em>Weekly Report</em> and <em>Summary</em> sheets.
                        </p>
                        <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                            <a href="{{ url_for('download_preview_report', date=selected_date.isoformat()) }}" class="btn btn-outline">
                                📥 Download Preview Excel
                            </a>
                            <form method="POST" action="{{ url_for('send_report_now') }}" style="display: inline;">
                                <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
                                <button type="submit" class="btn btn-primary" onclick="return confirm('Send weekly Excel report now to manager?');">
                                    ✉️ Send Weekly Report Now
                                </button>
                            </form>
                        </div>
                    </div>
                </div>

                <!-- Right Column: Semester Timetable Master & Live Settings -->
                <div>
                    <!-- Master Semester Timetable Card -->
                    <div class="card">
                        <div class="card-header">
                            <span class="card-title">🏫 Semester Timetable (Recurring)</span>
                        </div>
                        <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 12px;">
                            This schedule remains fixed for the entire semester and automatically populates daily schedule dispatches.
                        </p>

                        <!-- Extracted Picture Present Directly on Dashboard -->
                        {% if latest_img %}
                            <div style="margin-bottom: 16px; border: 1px solid var(--border); border-radius: 8px; padding: 10px; background: var(--bg-card-subtle); text-align: center;">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                                    <span style="font-weight: 600; font-size: 0.85rem; color: var(--text-main);">📸 Uploaded Timetable Picture</span>
                                    <a href="{{ url_for('preview_timetable') }}" target="_blank" class="btn btn-outline btn-sm">🔍 Open Full Image</a>
                                </div>
                                <img src="{{ url_for('preview_timetable') }}" alt="Uploaded Timetable Screenshot" style="width: 100%; max-height: 380px; object-fit: contain; border-radius: 6px; border: 1px solid var(--border); background: white;">
                                <p style="font-size: 0.75rem; color: var(--text-muted); margin-top: 6px;">
                                    File: {{ latest_img.filename }} • Saved in Database
                                </p>
                            </div>
                        {% endif %}

                        <!-- Day Tabs (Monday to Sunday) -->
                        <div class="tab-nav">
                            <button type="button" class="tab-btn active" onclick="switchDayTab(0, this)">Mon</button>
                            <button type="button" class="tab-btn" onclick="switchDayTab(1, this)">Tue</button>
                            <button type="button" class="tab-btn" onclick="switchDayTab(2, this)">Wed</button>
                            <button type="button" class="tab-btn" onclick="switchDayTab(3, this)">Thu</button>
                            <button type="button" class="tab-btn" onclick="switchDayTab(4, this)">Fri</button>
                            <button type="button" class="tab-btn" onclick="switchDayTab(5, this)">Sat</button>
                            <button type="button" class="tab-btn" onclick="switchDayTab(6, this)">Sun</button>
                        </div>

                        <!-- Slots Display per Day -->
                        {% for day_idx in range(7) %}
                            <div class="day-slot-list" id="day-slots-{{ day_idx }}" style="display: {% if day_idx == 0 %}block{% else %}none{% endif %}; margin-bottom: 14px;">
                                {% set day_slots = slots_by_day[day_idx] %}
                                {% if day_slots %}
                                    {% for s in day_slots %}
                                        <div class="schedule-item">
                                            <div class="schedule-time">🕒 {{ s.start_time }} – {{ s.end_time }}</div>
                                            <div class="schedule-detail">
                                                <div class="schedule-subject">{{ s.subject }}</div>
                                                <div class="schedule-meta">Room/Sec: {{ s.room or '—' }} | Faculty: {{ s.faculty or '—' }}</div>
                                            </div>
                                            <form method="POST" action="{{ url_for('delete_timetable_slot', slot_id=s.id) }}" style="display:inline;">
                                                <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
                                                <button type="submit" class="btn btn-outline btn-sm" title="Delete class" onclick="return confirm('Delete this class?');">✕</button>
                                            </form>
                                        </div>
                                    {% endfor %}
                                {% else %}
                                    <div style="font-size: 0.85rem; color: var(--text-muted); padding: 12px 0; text-align: center;">
                                        No recurring classes for this day. (You can add classes or use Manual Entry for weekend sessions).
                                    </div>
                                {% endif %}
                            </div>
                        {% endfor %}

                        <!-- Quick Add Class Form -->
                        <details style="margin-top: 10px; background: var(--bg-card-subtle); padding: 10px; border-radius: 6px;">
                            <summary style="font-size: 0.85rem; font-weight: 600; cursor: pointer; color: var(--primary);">
                                ➕ Add a recurring class to semester timetable
                            </summary>
                            <form method="POST" action="{{ url_for('add_timetable_slot') }}" style="margin-top: 10px;">
                                <input type="hidden" name="csrf_token" value="{{ csrf_token }}">

                                <div class="form-group">
                                    <label class="form-label">Day of Week</label>
                                    <select name="day_of_week" id="addSlotDay" class="form-control">
                                        <option value="0">Monday</option>
                                        <option value="1">Tuesday</option>
                                        <option value="2">Wednesday</option>
                                        <option value="3">Thursday</option>
                                        <option value="4">Friday</option>
                                        <option value="5">Saturday</option>
                                        <option value="6">Sunday</option>
                                    </select>
                                </div>

                                <div class="form-row">
                                    <div class="form-group">
                                        <label class="form-label">Start Time</label>
                                        <input type="text" name="start_time" class="form-control" placeholder="09:00 AM" required>
                                    </div>
                                    <div class="form-group">
                                        <label class="form-label">End Time</label>
                                        <input type="text" name="end_time" class="form-control" placeholder="10:00 AM" required>
                                    </div>
                                </div>

                                <div class="form-group">
                                    <label class="form-label">Subject / Course Name</label>
                                    <input type="text" name="subject" class="form-control" placeholder="E.g. Data Mining and Warehousing (DMW)" required>
                                </div>

                                <div class="form-row">
                                    <div class="form-group">
                                        <label class="form-label">Room / Section</label>
                                        <input type="text" name="room" class="form-control" placeholder="MCA B">
                                    </div>
                                    <div class="form-group">
                                        <label class="form-label">Faculty</label>
                                        <input type="text" name="faculty" class="form-control" placeholder="Prof. Name">
                                    </div>
                                </div>

                                <button type="submit" class="btn btn-primary btn-sm btn-block">Save to Semester Timetable</button>
                            </form>
                        </details>

                        <!-- Live Cloud Excel / Google Sheets Auto-Sync -->
                        <div style="margin-top: 16px; border-top: 1px solid var(--border); padding-top: 14px; background: var(--bg-card-subtle); padding: 12px; border-radius: 8px; margin-bottom: 14px;">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                                <strong style="font-size: 0.85rem; color: var(--primary);">🔗 Live Cloud Excel / Google Sheets Auto-Sync</strong>
                                {% if settings.timetable_excel_url %}
                                    <span class="badge badge-success">Auto-Sync Active</span>
                                {% else %}
                                    <span class="badge badge-warning">Not Connected</span>
                                {% endif %}
                            </div>
                            <p style="font-size: 0.75rem; color: var(--text-muted); margin-bottom: 8px;">
                                Paste your Google Sheets or OneDrive link once. The app will automatically sync latest classes before sending every night — <strong>no manual uploading needed!</strong>
                            </p>
                            <form method="POST" action="{{ url_for('sync_timetable_url') }}">
                                <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
                                <div style="display: flex; gap: 8px; flex-wrap: wrap;">
                                    <input type="url" name="timetable_excel_url" class="form-control" value="{{ settings.timetable_excel_url }}" placeholder="https://docs.google.com/spreadsheets/d/.../edit" style="flex: 1; min-width: 220px; font-size: 0.8rem;">
                                    <button type="submit" class="btn btn-primary btn-sm" style="white-space: nowrap;">
                                        🔄 Sync from Cloud
                                    </button>
                                </div>
                            </form>
                        </div>

                        <!-- Timetable Excel / File Upload & API -->
                        <div style="margin-top: 10px; border-top: 1px solid var(--border); padding-top: 14px;">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                                <strong style="font-size: 0.85rem; color: var(--text-main);">📥 Manual Import (Excel File / CSV / Screenshot)</strong>
                                <a href="{{ url_for('download_timetable_excel') }}" class="btn btn-outline btn-sm">📥 Download Excel Sheet</a>
                            </div>
                            <p style="font-size: 0.75rem; color: var(--text-muted); margin-bottom: 10px;">
                                Upload an Excel (.xlsx), CSV, or screenshot. Automatically parses class names, timings, and expands abbreviations.
                            </p>
                            <form method="POST" action="{{ url_for('upload_timetable') }}" enctype="multipart/form-data">
                                <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
                                <div style="display: flex; gap: 8px; flex-wrap: wrap;">
                                    <input type="file" name="timetable_file" class="form-control" accept=".xlsx,.xls,.csv,.png,.jpg,.jpeg,.webp" style="flex: 1; min-width: 200px;" required>
                                    <button type="submit" class="btn btn-outline btn-sm" style="white-space: nowrap;">Import File</button>
                                </div>
                            </form>
                            <div style="font-size: 0.7rem; color: var(--text-muted); margin-top: 6px;">
                                💡 REST API: <code>POST /api/timetable/upload</code> (file upload) • <code>GET /api/timetable/slots</code> (JSON)
                            </div>
                        </div>
                    </div>

                    <!-- Live App Settings Card (Always visible and easily editable) -->
                    <div class="card">
                        <div class="card-header">
                            <span class="card-title">⚙️ Email & App Settings</span>
                        </div>
                        <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 12px;">
                            Edit and save your email settings here. Any changes take effect immediately.
                        </p>

                        <form method="POST" action="{{ url_for('save_settings') }}">
                            <input type="hidden" name="csrf_token" value="{{ csrf_token }}">

                            <div class="form-group">
                                <label class="form-label" style="font-weight: 700; color: var(--primary);">
                                    Manager's Email Address (Default Recipient)
                                </label>
                                <input type="email" name="manager_email" class="form-control" value="{{ settings.manager_email }}" style="font-weight: 600;" required>
                            </div>

                            <div class="form-group">
                                <label class="form-label">Your Email Address (CC & Alerts)</label>
                                <input type="email" name="owner_email" class="form-control" value="{{ settings.owner_email }}" required>
                            </div>

                            <div class="form-row">
                                <div class="form-group">
                                    <label class="form-label">Schedule Send Time (Day Before)</label>
                                    <input type="text" name="schedule_time" class="form-control" value="{{ settings.schedule_time }}" placeholder="18:00">
                                </div>
                                <div class="form-group">
                                    <label class="form-label">Weekly Report Time</label>
                                    <input type="text" name="report_time" class="form-control" value="{{ settings.report_time }}" placeholder="17:00">
                                </div>
                            </div>

                            <div class="form-row">
                                <div class="form-group">
                                    <label class="form-label">Timezone</label>
                                    <input type="text" name="timezone" class="form-control" value="{{ settings.timezone }}">
                                </div>
                                <div class="form-group">
                                    <label class="form-label">Gemini API Key (Optional)</label>
                                    <input type="password" name="gemini_api_key" class="form-control" placeholder="Paste AI Studio Key">
                                </div>
                            </div>

                            <div class="form-group">
                                <label class="form-label">Cloud Excel / Google Sheets Sync URL (Optional)</label>
                                <input type="url" name="timetable_excel_url" class="form-control" value="{{ settings.timetable_excel_url }}" placeholder="https://docs.google.com/spreadsheets/d/.../edit">
                                <small style="font-size: 0.75rem; color: var(--text-muted); display: block; margin-top: 3px;">
                                    💡 Automatically pulls fresh classes from Google Sheets or OneDrive before sending every night.
                                </small>
                            </div>

                            <button type="submit" class="btn btn-primary btn-block btn-sm" style="margin-top: 8px;">
                                💾 Save Email & Settings
                            </button>
                        </form>

                        <div style="margin-top: 14px; padding-top: 10px; border-top: 1px solid var(--border); font-size: 0.75rem; color: var(--text-muted); line-height: 1.5;">
                            <div>🕒 Current Server Time: <strong>{{ now_str }}</strong></div>
                            <div>🤖 AI Parser: <strong>{% if settings.gemini_active %}{{ settings.gemini_model }}{% else %}Rule-Based Fallback{% endif %}</strong></div>
                            <div>📧 Gmail SMTP: <strong>{% if settings.smtp_configured %}Configured (Ready to Send){% else %}Missing credentials in .env{% endif %}</strong></div>
                        </div>
                    </div>
                </div>

            </div>

        {% endif %}

    </main>

    <footer class="footer">
        Work Reporter • Automated Daily Logs & Weekly Excel Reports
    </footer>

    <script>
        // Theme toggle
        (function() {
            const root = document.documentElement;
            const themeToggle = document.getElementById('themeToggle');
            const themeIcon = document.getElementById('themeIcon');
            const themeLabel = document.getElementById('themeLabel');

            function getPreferredTheme() {
                const stored = localStorage.getItem('wr_theme');
                if (stored) return stored;
                return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
            }

            function applyTheme(theme) {
                root.setAttribute('data-theme', theme);
                localStorage.setItem('wr_theme', theme);
                if (themeIcon && themeLabel) {
                    if (theme === 'dark') {
                        themeIcon.textContent = '☀️';
                        themeLabel.textContent = 'Light';
                    } else {
                        themeIcon.textContent = '🌙';
                        themeLabel.textContent = 'Dark';
                    }
                }
            }

            applyTheme(getPreferredTheme());

            if (themeToggle) {
                themeToggle.addEventListener('click', () => {
                    const current = root.getAttribute('data-theme') || 'light';
                    applyTheme(current === 'dark' ? 'light' : 'dark');
                });
            }
        })();

        // Date selection
        function selectDateCard(isoDate, noteContent) {
            const datePicker = document.getElementById('datePicker');
            const noteTextarea = document.getElementById('noteContent');
            if (datePicker && noteTextarea) {
                datePicker.value = isoDate;
                noteTextarea.value = noteContent || '';
            }
        }

        function handleDateChange(newDate) {
            if (newDate) {
                window.location.href = "{{ url_for('dashboard') }}?date=" + newDate;
            }
        }

        // Override Date Helper (Tomorrow, Saturday, Sunday)
        function setOverrideDate(dStr) {
            const input = document.getElementById('overrideDateInput');
            if (input) {
                input.value = dStr;
            }
        }

        // Semester Timetable Day Switcher (Mon=0 to Sun=6)
        function switchDayTab(dayIdx, btn) {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            for (let i = 0; i < 7; i++) {
                const el = document.getElementById('day-slots-' + i);
                if (el) el.style.display = (i === dayIdx) ? 'block' : 'none';
            }

            const select = document.getElementById('addSlotDay');
            if (select) select.value = dayIdx;
        }

        // Toggle Settings Edit Form
        function toggleSettingsEdit() {
            const view = document.getElementById('settingsView');
            const form = document.getElementById('settingsEditForm');
            if (view && form) {
                if (form.style.display === 'none') {
                    form.style.display = 'block';
                    view.style.display = 'none';
                } else {
                    form.style.display = 'none';
                    view.style.display = 'block';
                }
            }
        }
    </script>
</body>
</html>
"""

# ==============================================================================
# Application Entry Point
# ==============================================================================
if __name__ == "__main__":
    port = int(get_config("PORT", 5000))
    debug = get_config("FLASK_DEBUG", "false").lower() in ("1", "true", "yes")
    app.run(host="0.0.0.0", port=port, debug=debug)
