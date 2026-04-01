import io
import os
import sqlite3
from datetime import datetime, timezone
from urllib.parse import urlparse

from flask import Flask, request, jsonify, Response, render_template_string, send_file
from twilio.twiml.messaging_response import MessagingResponse
from openpyxl import Workbook

try:
    import psycopg
    from psycopg.rows import dict_row
except Exception:
    psycopg = None
    dict_row = None

APP_NAME = "WhatsApp Clock Agent"
DB_PATH = os.getenv("DB_PATH", "timeclock.db")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
ADMIN_NUMBERS = {
    n.strip() for n in os.getenv("ADMIN_NUMBERS", "").split(",") if n.strip()
}
TIMEZONE_LABEL = os.getenv("TIMEZONE_LABEL", "America/Denver")
ADMIN_TOKEN = os.getenv("ADMIN_TOKEN", "change-this-admin-token")

app = Flask(__name__)


# -----------------------------
# Database helpers
# -----------------------------
def using_postgres() -> bool:
    return DATABASE_URL.startswith("postgres://") or DATABASE_URL.startswith("postgresql://")


def normalized_database_url() -> str:
    if DATABASE_URL.startswith("postgres://"):
        return DATABASE_URL.replace("postgres://", "postgresql://", 1)
    return DATABASE_URL


def get_conn():
    if using_postgres():
        if psycopg is None:
            raise RuntimeError("psycopg is not installed. Add 'psycopg[binary]' to requirements.txt")
        return psycopg.connect(normalized_database_url(), row_factory=dict_row)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _convert_placeholders(sql: str) -> str:
    if using_postgres():
        return sql.replace("?", "%s")
    return sql


def db_execute(sql: str, params=(), fetchone=False, fetchall=False, commit=False):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(_convert_placeholders(sql), params)

    result = None
    if fetchone:
        result = cur.fetchone()
    elif fetchall:
        result = cur.fetchall()

    if commit:
        conn.commit()

    cur.close()
    conn.close()
    return result


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    if using_postgres():
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS employees (
                id BIGSERIAL PRIMARY KEY,
                phone TEXT UNIQUE NOT NULL,
                name TEXT
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS shifts (
                id BIGSERIAL PRIMARY KEY,
                phone TEXT NOT NULL,
                employee_name TEXT,
                date_local TEXT,
                clock_in_utc TEXT,
                clock_out_utc TEXT,
                lunch_minutes INTEGER DEFAULT 0,
                in_lat DOUBLE PRECISION,
                in_lng DOUBLE PRECISION,
                out_lat DOUBLE PRECISION,
                out_lng DOUBLE PRECISION,
                location_description_in TEXT,
                location_description_out TEXT,
                notes TEXT,
                total_work_minutes INTEGER,
                status TEXT DEFAULT 'open'
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS conversation_state (
                phone TEXT PRIMARY KEY,
                state TEXT,
                temp_value TEXT,
                updated_at_utc TEXT
            )
            """
        )
    else:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                phone TEXT UNIQUE NOT NULL,
                name TEXT
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS shifts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                phone TEXT NOT NULL,
                employee_name TEXT,
                date_local TEXT,
                clock_in_utc TEXT,
                clock_out_utc TEXT,
                lunch_minutes INTEGER DEFAULT 0,
                in_lat REAL,
                in_lng REAL,
                out_lat REAL,
                out_lng REAL,
                location_description_in TEXT,
                location_description_out TEXT,
                notes TEXT,
                total_work_minutes INTEGER,
                status TEXT DEFAULT 'open'
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS conversation_state (
                phone TEXT PRIMARY KEY,
                state TEXT,
                temp_value TEXT,
                updated_at_utc TEXT
            )
            """
        )

    conn.commit()
    cur.close()
    conn.close()


# -----------------------------
# Core utilities
# -----------------------------
def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def local_date_string() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def normalize_text(value: str) -> str:
    return (value or "").strip()


def parse_iso(value: str):
    if not value:
        return None
    return datetime.fromisoformat(value)


def fmt_dt(iso_value: str) -> str:
    dt = parse_iso(iso_value)
    if not dt:
        return "-"
    return dt.strftime("%Y-%m-%d %I:%M %p UTC")


def fmt_minutes(total_minutes: int) -> str:
    total_minutes = int(total_minutes or 0)
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}h {minutes}m"


def admin_authorized(req) -> bool:
    token = req.args.get("token") or req.headers.get("X-Admin-Token", "")
    return bool(ADMIN_TOKEN) and token == ADMIN_TOKEN


def set_state(phone: str, state: str, temp_value: str = ""):
    db_execute(
        """
        INSERT INTO conversation_state(phone, state, temp_value, updated_at_utc)
        VALUES(?, ?, ?, ?)
        ON CONFLICT(phone) DO UPDATE SET
            state = excluded.state,
            temp_value = excluded.temp_value,
            updated_at_utc = excluded.updated_at_utc
        """,
        (phone, state, temp_value, utc_now_iso()),
        commit=True,
    )


def get_state(phone: str):
    return db_execute(
        "SELECT * FROM conversation_state WHERE phone = ?",
        (phone,),
        fetchone=True,
    )


def clear_state(phone: str):
    db_execute("DELETE FROM conversation_state WHERE phone = ?", (phone,), commit=True)


def get_employee_name(phone: str):
    row = db_execute(
        "SELECT name FROM employees WHERE phone = ?",
        (phone,),
        fetchone=True,
    )
    return row["name"] if row else None


def set_employee_name(phone: str, name: str):
    db_execute(
        """
        INSERT INTO employees(phone, name)
        VALUES(?, ?)
        ON CONFLICT(phone) DO UPDATE SET name = excluded.name
        """,
        (phone, name),
        commit=True,
    )


def get_open_shift(phone: str):
    return db_execute(
        "SELECT * FROM shifts WHERE phone = ? AND status = 'open' ORDER BY id DESC LIMIT 1",
        (phone,),
        fetchone=True,
    )


def create_shift(phone: str, employee_name: str, loc_description: str = "", lat=None, lng=None):
    db_execute(
        """
        INSERT INTO shifts(
            phone, employee_name, date_local, clock_in_utc,
            in_lat, in_lng, location_description_in, status
        )
        VALUES(?, ?, ?, ?, ?, ?, ?, 'open')
        """,
        (
            phone,
            employee_name,
            local_date_string(),
            utc_now_iso(),
            lat,
            lng,
            loc_description,
        ),
        commit=True,
    )


def close_shift(phone: str, lunch_minutes: int = 0, notes: str = "", loc_description: str = "", lat=None, lng=None):
    open_shift = get_open_shift(phone)
    if not open_shift:
        return None

    clock_in = datetime.fromisoformat(open_shift["clock_in_utc"])
    clock_out = datetime.fromisoformat(utc_now_iso())
    total_minutes = int((clock_out - clock_in).total_seconds() // 60) - int(lunch_minutes)
    total_minutes = max(total_minutes, 0)

    db_execute(
        """
        UPDATE shifts
        SET clock_out_utc = ?,
            lunch_minutes = ?,
            out_lat = ?,
            out_lng = ?,
            location_description_out = ?,
            notes = ?,
            total_work_minutes = ?,
            status = 'closed'
        WHERE id = ?
        """,
        (
            clock_out.isoformat(),
            int(lunch_minutes),
            lat,
            lng,
            loc_description,
            notes,
            total_minutes,
            open_shift["id"],
        ),
        commit=True,
    )
    return total_minutes


def latest_closed_shifts(limit: int = 20):
    return db_execute(
        """
        SELECT employee_name, phone, date_local, lunch_minutes, total_work_minutes,
               location_description_in, location_description_out,
               clock_in_utc, clock_out_utc, notes
        FROM shifts
        WHERE status = 'closed'
        ORDER BY id DESC
        LIMIT ?
        """,
        (limit,),
        fetchall=True,
    )


def fetch_dashboard_shifts(employee: str = "", date_from: str = "", date_to: str = ""):
    sql = """
        SELECT id, employee_name, phone, date_local, clock_in_utc, clock_out_utc,
               lunch_minutes, location_description_in, location_description_out,
               in_lat, in_lng, out_lat, out_lng, notes, total_work_minutes, status
        FROM shifts
        WHERE 1=1
    """
    params = []

    if employee:
        sql += " AND (employee_name LIKE ? OR phone LIKE ?)"
        like = f"%{employee}%"
        params.extend([like, like])

    if date_from:
        sql += " AND date_local >= ?"
        params.append(date_from)

    if date_to:
        sql += " AND date_local <= ?"
        params.append(date_to)

    sql += " ORDER BY date_local DESC, id DESC"
    return db_execute(sql, tuple(params), fetchall=True)


def build_dashboard_summary(rows):
    total_minutes = sum(int(row["total_work_minutes"] or 0) for row in rows if row["status"] == "closed")
    unique_employees = len({(row["employee_name"] or row["phone"]) for row in rows})
    open_shifts = sum(1 for row in rows if row["status"] == "open")
    closed_shifts = sum(1 for row in rows if row["status"] == "closed")
    return {
        "total_minutes": total_minutes,
        "unique_employees": unique_employees,
        "open_shifts": open_shifts,
        "closed_shifts": closed_shifts,
    }


def build_excel(rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Shifts"

    headers = [
        "Employee",
        "Phone",
        "Date",
        "Clock In UTC",
        "Clock Out UTC",
        "Lunch Minutes",
        "Total Worked Minutes",
        "Total Worked",
        "Status",
        "Clock In Site",
        "Clock Out Site",
        "In Latitude",
        "In Longitude",
        "Out Latitude",
        "Out Longitude",
        "Notes",
    ]
    sheet.append(headers)

    for row in rows:
        sheet.append([
            row["employee_name"],
            row["phone"],
            row["date_local"],
            row["clock_in_utc"],
            row["clock_out_utc"],
            row["lunch_minutes"],
            row["total_work_minutes"],
            fmt_minutes(row["total_work_minutes"] or 0),
            row["status"],
            row["location_description_in"],
            row["location_description_out"],
            row["in_lat"],
            row["in_lng"],
            row["out_lat"],
            row["out_lng"],
            row["notes"],
        ])

    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        sheet.column_dimensions[column_letter].width = min(max_length + 2, 28)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


# -----------------------------
# Message handling
# -----------------------------
def help_text() -> str:
    return (
        "*Clock Agent Commands*\n\n"
        "start - register your name\n"
        "in - clock in\n"
        "out - clock out\n"
        "lunch 30 - set lunch minutes\n"
        "site Warehouse A - add work location description\n"
        "note Worked on delivery route - add notes before clock out\n"
        "status - view current open shift\n"
        "help - show commands\n\n"
        "Tip: for GPS, send a WhatsApp location message right after 'in' or right before 'out'."
    )


def parse_whatsapp_location(form):
    lat = form.get("Latitude") or form.get("WaLatitude")
    lng = form.get("Longitude") or form.get("WaLongitude")

    try:
        if lat is not None and lng is not None:
            return float(lat), float(lng)
    except ValueError:
        return None, None
    return None, None


def incoming_text(form) -> str:
    return normalize_text(form.get("Body", ""))


def from_number(form) -> str:
    return normalize_text(form.get("From", ""))


def is_location_message(form) -> bool:
    lat, lng = parse_whatsapp_location(form)
    return lat is not None and lng is not None


def save_location_to_open_shift(phone: str, lat: float, lng: float):
    shift = get_open_shift(phone)
    if not shift:
        return False, "No open shift found. Send *in* first."

    if shift["in_lat"] is None or shift["in_lng"] is None:
        db_execute(
            "UPDATE shifts SET in_lat = ?, in_lng = ? WHERE id = ?",
            (lat, lng, shift["id"]),
            commit=True,
        )
        return True, "Location saved for *clock in*."

    db_execute(
        "UPDATE shifts SET out_lat = ?, out_lng = ? WHERE id = ?",
        (lat, lng, shift["id"]),
        commit=True,
    )
    return True, "Location saved for current shift."


def handle_stateful_reply(phone: str, text: str):
    manual_out_lunch_reply = handle_manual_out_state(phone, text)
    if manual_out_lunch_reply:
        return manual_out_lunch_reply

    manual_out_description_reply = handle_manual_out_description_state(phone, text)
    if manual_out_description_reply:
        return manual_out_description_reply

    state = get_state(phone)
    if not state:
        return None

    current = state["state"]

    if current == "awaiting_name":
        name = text.strip()
        if not name:
            return "Send your full name to complete registration."
        set_employee_name(phone, name)
        clear_state(phone)
        return f"Perfect. You are registered as *{name}*. Now send *in* to clock in."

    if current == "awaiting_in_description":
        name = get_employee_name(phone) or "Employee"
        create_shift(phone, name, loc_description=text)
        clear_state(phone)
        return (
            f"✅ Clock in saved for *{name}*."
            f"Site: {text or 'Not provided'}"
            "Now send your WhatsApp location if you want GPS attached."
        )

    if current == "awaiting_out_lunch":
        try:
            lunch_minutes = int(text)
            if lunch_minutes < 0 or lunch_minutes > 240:
                return "Send lunch as minutes only, for example: 30"
        except ValueError:
            return "Send lunch as minutes only, for example: 30"

        set_state(phone, "awaiting_out_description", str(lunch_minutes))
        return "Got it. Now send the location description for clock out. Example: *Warehouse B*"

    if current == "awaiting_out_description":
        try:
            lunch_minutes = int(state["temp_value"] or "0")
        except ValueError:
            lunch_minutes = 0
        total = close_shift(phone, lunch_minutes=lunch_minutes, loc_description=text)
        clear_state(phone)
        if total is None:
            return "No open shift found. Send *in* first."
        return (
            f"✅ Clock out saved."
            f"Lunch: {lunch_minutes} min"
            f"Site: {text or 'Not provided'}"
            f"Worked: *{fmt_minutes(total)}*"
        )

    return None

    current = state["state"]

    if current == "awaiting_name":
        name = text.strip()
        if not name:
            return "Send your full name to complete registration."
        set_employee_name(phone, name)
        clear_state(phone)
        return f"Perfect. You are registered as *{name}*. Now send *in* to clock in."

    if current == "awaiting_in_description":
        name = get_employee_name(phone) or "Employee"
        create_shift(phone, name, loc_description=text)
        clear_state(phone)
        return (
            f"✅ Clock in saved for *{name}*.\n"
            f"Site: {text or 'Not provided'}\n"
            "Now send your WhatsApp location if you want GPS attached."
        )

    if current == "awaiting_out_lunch":
        try:
            lunch_minutes = int(text)
            if lunch_minutes < 0 or lunch_minutes > 240:
                return "Send lunch as minutes only, for example: 30"
        except ValueError:
            return "Send lunch as minutes only, for example: 30"

        set_state(phone, "awaiting_out_description", str(lunch_minutes))
        return "Got it. Now send the location description for clock out. Example: *Warehouse B*"

    if current == "awaiting_out_description":
        try:
            lunch_minutes = int(state["temp_value"] or "0")
        except ValueError:
            lunch_minutes = 0
        total = close_shift(phone, lunch_minutes=lunch_minutes, loc_description=text)
        clear_state(phone)
        if total is None:
            return "No open shift found. Send *in* first."
        return (
            f"✅ Clock out saved.\n"
            f"Lunch: {lunch_minutes} min\n"
            f"Site: {text or 'Not provided'}\n"
            f"Worked: *{fmt_minutes(total)}*"
        )

    return None


def parse_manual_time(raw: str, base_date: str):
    raw = normalize_text(raw).lower()
    if not raw:
        return None

    normalized = (
        raw.replace("a.m.", "am")
        .replace("p.m.", "pm")
        .replace("a.m", "am")
        .replace("p.m", "pm")
        .replace(" am", "am")
        .replace(" pm", "pm")
    )

    patterns = [
        "%H:%M",
        "%I:%M%p",
        "%I:%M %p",
        "%I%p",
        "%I %p",
        "%H%M",
    ]

    for pattern in patterns:
        try:
            parsed = datetime.strptime(normalized, pattern)
            base = datetime.strptime(base_date, "%Y-%m-%d")
            return base.replace(hour=parsed.hour, minute=parsed.minute, second=0, microsecond=0)
        except ValueError:
            continue
    return None


def extract_turno_payload(payload: str):
    original = normalize_text(payload)
    lowered = original.lower()

    normalized = lowered
    replacements = {
        "minutos": "min",
        "mins": "min",
        "minuto": "min",
        "min": "min",
        "lonch": "lonche",
        "lonce": "lonche",
        "lunch": "lonche",
        "lounche": "lonche",
        "lonchee": "lonche",
        "lonhce": "lonche",
    }
    for old, new in replacements.items():
        normalized = normalized.replace(old, new)

    normalized = normalized.replace("–", "-").replace("—", "-")
    tokens = normalized.split()

    if not tokens:
        return None

    site = ""
    lunch_minutes = None
    in_raw = None
    out_raw = None

    lonche_index = None
    for i, token in enumerate(tokens):
        if token == "lonche":
            lonche_index = i
            break

    if lonche_index is not None:
        before = tokens[:lonche_index]
        after = tokens[lonche_index + 1 :]
        site = " ".join(after).strip()
    else:
        before = tokens

    cleaned_before = []
    for token in before:
        if token.endswith("min") and token[:-3].isdigit():
            lunch_minutes = int(token[:-3])
        elif token.isdigit() and lunch_minutes is None and len(cleaned_before) >= 2:
            lunch_minutes = int(token)
        else:
            cleaned_before.append(token)

    before = cleaned_before

    if before:
        first = before[0]
        if "-" in first and not first.startswith("-") and not first.endswith("-"):
            left, right = first.split("-", 1)
            in_raw = left
            out_raw = right
            if len(before) > 1 and before[1] in {"am", "pm"}:
                in_raw += before[1]
                if len(before) > 2 and before[2] in {"am", "pm"}:
                    out_raw += before[2]
            elif len(before) > 1 and (before[1].endswith("am") or before[1].endswith("pm")):
                out_raw += before[1]
        elif len(before) >= 4:
            in_raw = " ".join(before[:2])
            out_raw = " ".join(before[2:4])
        elif len(before) >= 2 and "-" in " ".join(before[:2]):
            compact = " ".join(before[:2])
            left, right = compact.split("-", 1)
            in_raw = left.strip()
            out_raw = right.strip()

    if lunch_minutes is None:
        for token in tokens:
            digits = "".join(ch for ch in token if ch.isdigit())
            if digits and (token.endswith("min") or token == digits):
                lunch_minutes = int(digits)
                break

    if not in_raw or not out_raw or lunch_minutes is None:
        return None

    return {
        "in_raw": in_raw.strip(),
        "out_raw": out_raw.strip(),
        "lunch_minutes": lunch_minutes,
        "site": site,
    }

    patterns = [
        "%H:%M",
        "%I:%M%p",
        "%I:%M %p",
        "%I%p",
        "%I %p",
    ]

    for pattern in patterns:
        try:
            parsed = datetime.strptime(raw, pattern)
            base = datetime.strptime(base_date, "%Y-%m-%d")
            return base.replace(hour=parsed.hour, minute=parsed.minute, second=0, microsecond=0)
        except ValueError:
            continue
    return None


def create_shift_at_time(phone: str, employee_name: str, clock_in_dt, loc_description: str = "", lat=None, lng=None):
    db_execute(
        """
        INSERT INTO shifts(
            phone, employee_name, date_local, clock_in_utc,
            in_lat, in_lng, location_description_in, status
        )
        VALUES(?, ?, ?, ?, ?, ?, ?, 'open')
        """,
        (
            phone,
            employee_name,
            clock_in_dt.strftime("%Y-%m-%d"),
            clock_in_dt.replace(tzinfo=timezone.utc).isoformat(),
            lat,
            lng,
            loc_description,
        ),
        commit=True,
    )


def close_shift_at_time(phone: str, clock_out_dt, lunch_minutes: int = 0, notes: str = "", loc_description: str = "", lat=None, lng=None):
    open_shift = get_open_shift(phone)
    if not open_shift:
        return None

    clock_in = datetime.fromisoformat(open_shift["clock_in_utc"])
    if clock_in.tzinfo is not None:
        clock_in = clock_in.astimezone(timezone.utc).replace(tzinfo=None)

    total_minutes = int((clock_out_dt - clock_in).total_seconds() // 60) - int(lunch_minutes)
    total_minutes = max(total_minutes, 0)

    db_execute(
        """
        UPDATE shifts
        SET clock_out_utc = ?,
            lunch_minutes = ?,
            out_lat = ?,
            out_lng = ?,
            location_description_out = ?,
            notes = ?,
            total_work_minutes = ?,
            status = 'closed'
        WHERE id = ?
        """,
        (
            clock_out_dt.replace(tzinfo=timezone.utc).isoformat(),
            int(lunch_minutes),
            lat,
            lng,
            loc_description,
            notes,
            total_minutes,
            open_shift["id"],
        ),
        commit=True,
    )
    return total_minutes


def handle_command(phone: str, text: str):
    lower = text.lower().strip()
    name = get_employee_name(phone)

    if lower in {"help", "menu", "ayuda"}:
        return (
            help_text()
            + "

*Offline Pro*
"
            + "You can also send manual times if there was no signal:
"
            + "in 8:00 am
"
            + "out 5:30 pm
"
            + "latein 8:00 am Job Site 4
"
            + "lateout 5:30 pm 30 Warehouse B
"
            + "shift 8:00 am 5:30 pm 30 Warehouse B
"
            + "turno 8:00 am 5:30 pm 30 lonche Rancho
"
            + "turno 8:00am 5:30pm 30min lonche Rancho
"
            + "turno 8:00-5:30 30min lonche Rancho
"
            + "turno 8 530 30 lonche Rancho"
        )

    if lower == "start":
        set_state(phone, "awaiting_name")
        return "Send your full name to register this WhatsApp number."

    if lower == "status":
        shift = get_open_shift(phone)
        if not shift:
            return "You do not have an open shift. Send *in* to clock in."
        return (
            f"Open shift found for *{shift['employee_name'] or 'Employee'}*."
            f"Date: {shift['date_local']}"
            f"Clock in: {shift['clock_in_utc']} UTC"
            f"Site: {shift['location_description_in'] or 'Not provided'}"
        )

    if lower == "in":
        if not name:
            set_state(phone, "awaiting_name")
            return "First register your name. Send your full name."
        if get_open_shift(phone):
            return "You already have an open shift. Send *status* or *out*."
        set_state(phone, "awaiting_in_description")
        return "Send the work location description for clock in. Example: *Job Site 4*"

    if lower.startswith("in "):
        if not name:
            set_state(phone, "awaiting_name")
            return "First register your name. Send your full name."
        if get_open_shift(phone):
            return "You already have an open shift. Send *status* or *out*."
        manual_dt = parse_manual_time(text[3:].strip(), local_date_string())
        if not manual_dt:
            return "Use a valid manual time, for example: *in 8:00 am*"
        create_shift_at_time(phone, name, manual_dt)
        return f"✅ Manual clock in saved at *{manual_dt.strftime('%I:%M %p')}*. Send *site Warehouse A* if you want to add the location description."

    if lower.startswith("latein "):
        if not name:
            set_state(phone, "awaiting_name")
            return "First register your name. Send your full name."
        if get_open_shift(phone):
            return "You already have an open shift. Send *status* or *out*."
        payload = text[7:].strip()
        pieces = payload.split(" ")
        if len(pieces) < 2:
            return "Use: *latein 8:00 am Job Site 4*"
        time_guess = " ".join(pieces[:2])
        site = " ".join(pieces[2:]).strip()
        manual_dt = parse_manual_time(time_guess, local_date_string())
        if not manual_dt:
            return "Use: *latein 8:00 am Job Site 4*"
        create_shift_at_time(phone, name, manual_dt, loc_description=site)
        return f"✅ Offline clock in saved at *{manual_dt.strftime('%I:%M %p')}* for *{site or 'No site'}*."

    if lower == "out":
        if not get_open_shift(phone):
            return "No open shift found. Send *in* first."
        set_state(phone, "awaiting_out_lunch")
        return "Send lunch minutes only. Example: *30*"

    if lower.startswith("out "):
        if not get_open_shift(phone):
            return "No open shift found. Send *in* first."
        manual_dt = parse_manual_time(text[4:].strip(), local_date_string())
        if not manual_dt:
            return "Use a valid manual time, for example: *out 5:30 pm*"
        set_state(phone, "awaiting_manual_out_lunch", manual_dt.strftime("%Y-%m-%d %H:%M"))
        return "Manual clock out time saved. Now send lunch minutes only. Example: *30*"

    if lower.startswith("lateout "):
        if not get_open_shift(phone):
            return "No open shift found. Send *in* first."
        payload = text[8:].strip()
        parts = payload.split(" ")
        if len(parts) < 4:
            return "Use: *lateout 5:30 pm 30 Warehouse B*"
        time_guess = " ".join(parts[:2])
        lunch_guess = parts[2]
        site = " ".join(parts[3:]).strip()
        manual_dt = parse_manual_time(time_guess, local_date_string())
        if not manual_dt:
            return "Use: *lateout 5:30 pm 30 Warehouse B*"
        try:
            lunch_minutes = int(lunch_guess)
        except ValueError:
            return "Use: *lateout 5:30 pm 30 Warehouse B*"
        total = close_shift_at_time(phone, manual_dt, lunch_minutes=lunch_minutes, loc_description=site)
        if total is None:
            return "No open shift found. Send *in* first."
        return (
            f"✅ Offline clock out saved at *{manual_dt.strftime('%I:%M %p')}*."
            f"Lunch: {lunch_minutes} min"
            f"Site: {site or 'Not provided'}"
            f"Worked: *{fmt_minutes(total)}*"
        )

        if lower.startswith("shift "):
        if not name:
            set_state(phone, "awaiting_name")
            return "First register your name. Send your full name."
        if get_open_shift(phone):
            return "You already have an open shift. Send *status* or *out* before using *shift*."

        payload = text[6:].strip()
        parts = payload.split()
        if len(parts) < 6:
            return "Use: *shift 8:00 am 5:30 pm 30 Warehouse B*"

        in_time_guess = " ".join(parts[:2])
        out_time_guess = " ".join(parts[2:4])
        lunch_guess = parts[4]
        site = " ".join(parts[5:]).strip()

        in_dt = parse_manual_time(in_time_guess, local_date_string())
        out_dt = parse_manual_time(out_time_guess, local_date_string())
        if not in_dt or not out_dt:
            return "Use: *shift 8:00 am 5:30 pm 30 Warehouse B*"

        try:
            lunch_minutes = int(lunch_guess)
        except ValueError:
            return "Use: *shift 8:00 am 5:30 pm 30 Warehouse B*"

        if out_dt <= in_dt:
            return "Clock out time must be later than clock in time."

        create_shift_at_time(phone, name, in_dt, loc_description=site)
        total = close_shift_at_time(phone, out_dt, lunch_minutes=lunch_minutes, loc_description=site)
        if total is None:
            return "Could not save the shift. Try again."

        return (
            f"✅ Shift saved in one message."
            f"In: *{in_dt.strftime('%I:%M %p')}*"
            f"Out: *{out_dt.strftime('%I:%M %p')}*"
            f"Lunch: {lunch_minutes} min"
            f"Site: {site or 'Not provided'}"
            f"Worked: *{fmt_minutes(total)}*"
        )

        if lower.startswith("turno "):
        if not name:
            set_state(phone, "awaiting_name")
            return "First register your name. Send your full name."
        if get_open_shift(phone):
            return "You already have an open shift. Send *status* or *out* before using *turno*."

        parsed = extract_turno_payload(text[6:])
        if not parsed:
            return (
                "Usa algo como:"
                "*Turno 8:00 am 5:30 pm 30 lonche Rancho*"
                "*Turno 8:00am 5:30pm 30min lonche Rancho*"
                "*Turno 8:00-5:30 30min lonche Rancho*"
                "*Turno 8 530 30 lonche Rancho*"
            )

        in_dt = parse_manual_time(parsed["in_raw"], local_date_string())
        out_dt = parse_manual_time(parsed["out_raw"], local_date_string())
        lunch_minutes = parsed["lunch_minutes"]
        site = parsed["site"]

        if not in_dt or not out_dt:
            return (
                "No entendí bien la hora. Usa algo como:"
                "*Turno 8:00 am 5:30 pm 30 lonche Rancho*"
                "o *Turno 8:00-5:30 30min lonche Rancho*"
            )

        if out_dt <= in_dt:
            return "La hora de salida debe ser después de la entrada."

        create_shift_at_time(phone, name, in_dt, loc_description=site)
        total = close_shift_at_time(phone, out_dt, lunch_minutes=lunch_minutes, loc_description=site)
        if total is None:
            return "No pude guardar el turno. Intenta otra vez."

        return (
            f"✅ Turno guardado en un mensaje."
            f"Entrada: *{in_dt.strftime('%I:%M %p')}*\"
            f"Salida: *{out_dt.strftime('%I:%M %p')}*"
            f"Lonche: {lunch_minutes} min"
            f"Lugar: {site or 'No especificado'}"
            f"Total trabajado: *{fmt_minutes(total)}*"
        )

    if lower.startswith("lunch "):
        try:
            minutes = int(lower.split(" ", 1)[1].strip())
        except ValueError:
            return "Use: *lunch 30*"
        shift = get_open_shift(phone)
        if not shift:
            return "No open shift found. Send *in* first."
        db_execute(
            "UPDATE shifts SET lunch_minutes = ? WHERE id = ?",
            (minutes, shift["id"]),
            commit=True,
        )
        return f"Lunch updated to *{minutes}* minutes for your current shift."

    if lower.startswith("site "):
        site_text = text[5:].strip()
        shift = get_open_shift(phone)
        if not shift:
            return "No open shift found. Send *in* first."
        db_execute(
            "UPDATE shifts SET location_description_in = ? WHERE id = ?",
            (site_text, shift["id"]),
            commit=True,
        )
        return f"Site saved: *{site_text}*"

    if lower.startswith("note "):
        note_text = text[5:].strip()
        shift = get_open_shift(phone)
        if not shift:
            return "No open shift found. Send *in* first."
        db_execute(
            "UPDATE shifts SET notes = ? WHERE id = ?",
            (note_text, shift["id"]),
            commit=True,
        )
        return "Note saved to your current shift."

    if lower == "report":
        if phone not in ADMIN_NUMBERS:
            return "This command is only available for admin numbers."
        rows = latest_closed_shifts(limit=10)
        if not rows:
            return "No closed shifts yet."
        lines = ["*Latest closed shifts*
"]
        for row in rows:
            lines.append(
                f"• {row['employee_name'] or row['phone']} | {row['date_local']} | {fmt_minutes(row['total_work_minutes'] or 0)}"
            )
        return "
".join(lines)

    return (
        "I did not understand that."
        f"{help_text()}"
    )


def handle_manual_out_state(phone: str, text: str):(phone: str, text: str):
    state = get_state(phone)
    if not state or state["state"] != "awaiting_manual_out_lunch":
        return None

    try:
        lunch_minutes = int(text.strip())
    except ValueError:
        return "Send lunch as minutes only, for example: 30"

    manual_dt = datetime.strptime(state["temp_value"], "%Y-%m-%d %H:%M")
    set_state(phone, "awaiting_manual_out_description", f"{state['temp_value']}|{lunch_minutes}")
    return f"Got it. Now send the location description for manual clock out at *{manual_dt.strftime('%I:%M %p')}*."


def handle_manual_out_description_state(phone: str, text: str):
    state = get_state(phone)
    if not state or state["state"] != "awaiting_manual_out_description":
        return None

    raw = state["temp_value"] or ""
    try:
        dt_raw, lunch_raw = raw.split("|", 1)
        manual_dt = datetime.strptime(dt_raw, "%Y-%m-%d %H:%M")
        lunch_minutes = int(lunch_raw)
    except Exception:
        clear_state(phone)
        return "Manual clock out data expired. Please send *out 5:30 pm* again."

    total = close_shift_at_time(phone, manual_dt, lunch_minutes=lunch_minutes, loc_description=text)
    clear_state(phone)
    if total is None:
        return "No open shift found. Send *in* first."

    return (
        f"✅ Manual clock out saved at *{manual_dt.strftime('%I:%M %p')}*."
        f"Lunch: {lunch_minutes} min"
        f"Site: {text or 'Not provided'}"
        f"Worked: *{fmt_minutes(total)}*"
    )


# -----------------------------
# Web routes
# -----------------------------
@app.route("/")
def home():
    return jsonify({
        "app": APP_NAME,
        "ok": True,
        "database": "postgres" if using_postgres() else "sqlite",
        "message": "Use /health, /dashboard?token=YOUR_TOKEN or /export.xlsx?token=YOUR_TOKEN",
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "ok": True,
        "app": APP_NAME,
        "timezone": TIMEZONE_LABEL,
        "database": "postgres" if using_postgres() else "sqlite",
    })


@app.route("/dashboard", methods=["GET"])
def dashboard():
    if not admin_authorized(request):
        return Response("Unauthorized", status=401)

    employee = normalize_text(request.args.get("employee", ""))
    date_from = normalize_text(request.args.get("date_from", ""))
    date_to = normalize_text(request.args.get("date_to", ""))
    rows = fetch_dashboard_shifts(employee=employee, date_from=date_from, date_to=date_to)
    summary = build_dashboard_summary(rows)

    html = """
    <!doctype html>
    <html lang="en">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Clock Agent Dashboard</title>
        <style>
            * { box-sizing: border-box; }
            body { margin: 0; font-family: Arial, sans-serif; background: #0b1020; color: #eef2ff; }
            .wrap { max-width: 1400px; margin: 0 auto; padding: 24px; }
            .title { font-size: 32px; font-weight: 700; margin-bottom: 6px; }
            .sub { color: #b7c0d8; margin-bottom: 24px; }
            .cards { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 16px; margin-bottom: 24px; }
            .card { background: #121933; border: 1px solid #283255; border-radius: 16px; padding: 18px; }
            .label { color: #9fb0db; font-size: 13px; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.05em; }
            .value { font-size: 28px; font-weight: 700; }
            .panel { background: #121933; border: 1px solid #283255; border-radius: 18px; padding: 18px; margin-bottom: 20px; }
            form { display: grid; grid-template-columns: 1.2fr 1fr 1fr auto auto; gap: 12px; align-items: end; }
            label { display: block; font-size: 13px; color: #9fb0db; margin-bottom: 6px; }
            input { width: 100%; padding: 12px 14px; border-radius: 12px; border: 1px solid #33416f; background: #0f1530; color: #eef2ff; }
            .btn { display: inline-block; padding: 12px 16px; border-radius: 12px; border: 0; text-decoration: none; font-weight: 700; cursor: pointer; }
            .btn-primary { background: #6d7cff; color: white; }
            .btn-secondary { background: #1c2547; color: #eef2ff; border: 1px solid #33416f; }
            table { width: 100%; border-collapse: collapse; font-size: 14px; }
            th, td { padding: 12px; border-bottom: 1px solid #223055; text-align: left; vertical-align: top; }
            th { color: #9fb0db; font-size: 12px; text-transform: uppercase; letter-spacing: 0.04em; }
            .pill { display: inline-block; padding: 6px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; }
            .open { background: #583d00; color: #ffd977; }
            .closed { background: #0d4b32; color: #8df0bf; }
            .tools { display: flex; gap: 12px; margin-top: 14px; flex-wrap: wrap; }
            @media (max-width: 980px) { .cards { grid-template-columns: 1fr 1fr; } form { grid-template-columns: 1fr; } }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="title">WhatsApp Clock Dashboard</div>
            <div class="sub">See every shift, filter by employee or date, and export to Excel.</div>
            <div class="cards">
                <div class="card"><div class="label">Employees</div><div class="value">{{ summary.unique_employees }}</div></div>
                <div class="card"><div class="label">Closed Shifts</div><div class="value">{{ summary.closed_shifts }}</div></div>
                <div class="card"><div class="label">Open Shifts</div><div class="value">{{ summary.open_shifts }}</div></div>
                <div class="card"><div class="label">Worked Hours</div><div class="value">{{ worked_hours }}</div></div>
            </div>
            <div class="panel">
                <form method="get" action="/dashboard">
                    <input type="hidden" name="token" value="{{ token }}">
                    <div><label>Employee or phone</label><input name="employee" value="{{ employee }}" placeholder="Daniel or +1555..."></div>
                    <div><label>From date</label><input type="date" name="date_from" value="{{ date_from }}"></div>
                    <div><label>To date</label><input type="date" name="date_to" value="{{ date_to }}"></div>
                    <button class="btn btn-primary" type="submit">Filter</button>
                    <a class="btn btn-secondary" href="/dashboard?token={{ token }}">Reset</a>
                </form>
                <div class="tools">
                    <a class="btn btn-primary" href="/export.xlsx?token={{ token }}&employee={{ employee }}&date_from={{ date_from }}&date_to={{ date_to }}">Export Excel</a>
                </div>
            </div>
            <div class="panel">
                <table>
                    <thead>
                        <tr>
                            <th>Employee</th><th>Date</th><th>In</th><th>Out</th><th>Lunch</th><th>Total</th><th>Status</th><th>Sites</th><th>GPS</th><th>Notes</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in rows %}
                        <tr>
                            <td><strong>{{ row['employee_name'] or '-' }}</strong><br><span>{{ row['phone'] }}</span></td>
                            <td>{{ row['date_local'] or '-' }}</td>
                            <td>{{ fmt_dt(row['clock_in_utc']) }}</td>
                            <td>{{ fmt_dt(row['clock_out_utc']) }}</td>
                            <td>{{ row['lunch_minutes'] or 0 }} min</td>
                            <td>{{ fmt_minutes(row['total_work_minutes'] or 0) }}</td>
                            <td><span class="pill {{ 'open' if row['status'] == 'open' else 'closed' }}">{{ row['status'] }}</span></td>
                            <td><strong>In:</strong> {{ row['location_description_in'] or '-' }}<br><strong>Out:</strong> {{ row['location_description_out'] or '-' }}</td>
                            <td><strong>In:</strong> {{ row['in_lat'] or '-' }}, {{ row['in_lng'] or '-' }}<br><strong>Out:</strong> {{ row['out_lat'] or '-' }}, {{ row['out_lng'] or '-' }}</td>
                            <td>{{ row['notes'] or '-' }}</td>
                        </tr>
                        {% else %}
                        <tr><td colspan="10">No shifts found for the current filters.</td></tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </body>
    </html>
    """

    return render_template_string(
        html,
        rows=rows,
        summary=summary,
        fmt_minutes=fmt_minutes,
        fmt_dt=fmt_dt,
        worked_hours=fmt_minutes(summary["total_minutes"]),
        token=request.args.get("token", ""),
        employee=employee,
        date_from=date_from,
        date_to=date_to,
    )


@app.route("/export.xlsx", methods=["GET"])
def export_xlsx():
    if not admin_authorized(request):
        return Response("Unauthorized", status=401)

    employee = normalize_text(request.args.get("employee", ""))
    date_from = normalize_text(request.args.get("date_from", ""))
    date_to = normalize_text(request.args.get("date_to", ""))
    rows = fetch_dashboard_shifts(employee=employee, date_from=date_from, date_to=date_to)
    output = build_excel(rows)

    filename = f"clock-report-{datetime.now().strftime('%Y-%m-%d-%H%M')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/whatsapp", methods=["POST"])
def whatsapp_webhook():
    form = request.form
    phone = from_number(form)
    text = incoming_text(form)

    response = MessagingResponse()

    if is_location_message(form):
        lat, lng = parse_whatsapp_location(form)
        ok, msg = save_location_to_open_shift(phone, lat, lng)
        response.message(msg)
        return str(response)

    if text:
        state_reply = handle_stateful_reply(phone, text)
        if state_reply:
            response.message(state_reply)
            return str(response)

        reply = handle_command(phone, text)
        response.message(reply)
        return str(response)

    response.message("Send *help* to see available commands.")
    return str(response)


init_db()

if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
