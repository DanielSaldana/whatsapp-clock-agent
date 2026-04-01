import io
import os
import sqlite3
from datetime import datetime, timezone
from flask import Flask, request, jsonify, Response, render_template_string, send_file
from twilio.twiml.messaging_response import MessagingResponse
from openpyxl import Workbook

APP_NAME = "WhatsApp Clock Agent"
DB_PATH = os.getenv("DB_PATH", "timeclock.db")
ADMIN_NUMBERS = {
    n.strip() for n in os.getenv("ADMIN_NUMBERS", "").split(",") if n.strip()
}
TIMEZONE_LABEL = os.getenv("TIMEZONE_LABEL", "America/Denver")
ADMIN_TOKEN = os.getenv("ADMIN_TOKEN", "change-this-admin-token")

app = Flask(__name__)


# -----------------------------
# Database helpers
# -----------------------------
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()

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
    conn = get_conn()
    conn.execute(
        """
        INSERT INTO conversation_state(phone, state, temp_value, updated_at_utc)
        VALUES(?, ?, ?, ?)
        ON CONFLICT(phone) DO UPDATE SET
            state = excluded.state,
            temp_value = excluded.temp_value,
            updated_at_utc = excluded.updated_at_utc
        """,
        (phone, state, temp_value, utc_now_iso()),
    )
    conn.commit()
    conn.close()


def get_state(phone: str):
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM conversation_state WHERE phone = ?", (phone,)
    ).fetchone()
    conn.close()
    return row


def clear_state(phone: str):
    conn = get_conn()
    conn.execute("DELETE FROM conversation_state WHERE phone = ?", (phone,))
    conn.commit()
    conn.close()


def get_employee_name(phone: str):
    conn = get_conn()
    row = conn.execute(
        "SELECT name FROM employees WHERE phone = ?", (phone,)
    ).fetchone()
    conn.close()
    return row["name"] if row else None


def set_employee_name(phone: str, name: str):
    conn = get_conn()
    conn.execute(
        """
        INSERT INTO employees(phone, name)
        VALUES(?, ?)
        ON CONFLICT(phone) DO UPDATE SET name = excluded.name
        """,
        (phone, name),
    )
    conn.commit()
    conn.close()


def get_open_shift(phone: str):
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM shifts WHERE phone = ? AND status = 'open' ORDER BY id DESC LIMIT 1",
        (phone,),
    ).fetchone()
    conn.close()
    return row


def create_shift(phone: str, employee_name: str, loc_description: str = "", lat=None, lng=None):
    conn = get_conn()
    conn.execute(
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
    )
    conn.commit()
    conn.close()


def close_shift(phone: str, lunch_minutes: int = 0, notes: str = "", loc_description: str = "", lat=None, lng=None):
    open_shift = get_open_shift(phone)
    if not open_shift:
        return None

    clock_in = datetime.fromisoformat(open_shift["clock_in_utc"])
    clock_out = datetime.fromisoformat(utc_now_iso())
    total_minutes = int((clock_out - clock_in).total_seconds() // 60) - int(lunch_minutes)
    total_minutes = max(total_minutes, 0)

    conn = get_conn()
    conn.execute(
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
    )
    conn.commit()
    conn.close()
    return total_minutes


def latest_closed_shifts(limit: int = 20):
    conn = get_conn()
    rows = conn.execute(
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
    ).fetchall()
    conn.close()
    return rows


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

    conn = get_conn()
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return rows


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

# --- Simple turno parser (clean & minimal) ---
def parse_turno_simple(text: str):
    t = normalize_text(text).lower()

    # tolerate typos for 'turno'
    if not any(t.startswith(x) for x in ["turno ", "turn ", "trno ", "truno ", "tuno "]):
        return None

    payload = t.split(" ", 1)[1].strip() if " " in t else ""

    # normalize common variants / typos
    replacements = {
        "minutos": "min",
        "mins": "min",
        "minuto": "min",
        "min": "min",
        "lunch": "lonche",
        "lonch": "lonche",
        "lonce": "lonche",
        "lonhce": "lonche",
        "lonchee": "lonche",
        "lounche": "lonche",
    }
    for k, v in replacements.items():
        payload = payload.replace(k, v)

    payload = payload.replace("–", "-").replace("—", "-")

    # --- normalize time formats ---
    import re

    def normalize_time_token(token):
        token = token.strip()

        # 530 -> 5:30
        if token.isdigit() and len(token) in [3,4]:
            h = token[:-2]
            m = token[-2:]
            return str(int(h)) + ":" + m

        # 8am / 530pm
        match = re.match(r"^(\d{1,4})(am|pm)$", token)
        if match:
            num, ampm = match.groups()
            if len(num) in [3,4]:
                h = num[:-2]
                m = num[-2:]
                return str(int(h)) + ":" + m + " " + ampm
            return str(int(num)) + ":00 " + ampm

        return token

    # normalize dash times like 8-530 / 8-5:30 / 8am-5:30pm
    if "-" in payload:
        left, right = payload.split("-", 1)
        left = normalize_time_token(left.strip())
        right_parts = right.strip().split(" ", 1)
        right_time = normalize_time_token(right_parts[0])
        rest = right_parts[1] if len(right_parts) > 1 else ""
        payload = (left + " " + right_time + " " + rest).strip()
    t = normalize_text(text).lower()
    if not t.startswith("turno "):
        return None

    payload = t[6:].strip()

    # normalize common variants
    payload = payload.replace("minutos", "min").replace("mins", "min")
    payload = payload.replace("lunch", "lonche").replace("lonch", "lonche").replace("lonce", "lonche")

    # split site by 'lonche'
    if "lonche" not in payload:
        return None

    before, after = payload.split("lonche", 1)
    site = after.strip().title()

    parts = before.strip().split()

    # formats supported:
    # 8:00 am 5:30 pm 30
    # 8:00am 5:30pm 30
    # 8:00-5:30 30min

    try:
        if "-" in parts[0]:
            # format: 8:00-5:30
            start, end = parts[0].split("-")
            lunch = parts[1]
        else:
            # format: 8:00 am 5:30 pm 30
            start = parts[0] + (" " + parts[1] if "am" in parts[1] or "pm" in parts[1] else "")
            if "am" in parts[1] or "pm" in parts[1]:
                end = parts[2] + " " + parts[3]
                lunch = parts[4]
            else:
                end = parts[1]
                lunch = parts[2]

        lunch = int("".join([c for c in lunch if c.isdigit()]))

        return start.strip(), end.strip(), lunch, site
    except:
        return None

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

    conn = get_conn()
    if shift["in_lat"] is None or shift["in_lng"] is None:
        conn.execute(
            "UPDATE shifts SET in_lat = ?, in_lng = ? WHERE id = ?",
            (lat, lng, shift["id"]),
        )
        conn.commit()
        conn.close()
        return True, "Location saved for *clock in*."

    conn.execute(
        "UPDATE shifts SET out_lat = ?, out_lng = ? WHERE id = ?",
        (lat, lng, shift["id"]),
    )
    conn.commit()
    conn.close()
    return True, "Location saved for current shift."


def handle_stateful_reply(phone: str, text: str):
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


def handle_command(phone: str, text: str):
    # --- NEW: Turno single message ---
    parsed_turno = parse_turno_simple(text)
    if parsed_turno:
        name = get_employee_name(phone)
        if not name:
            set_state(phone, "awaiting_name")
            return "First register your name. Send your full name."

        if get_open_shift(phone):
            return "You already have an open shift. Send *out* first."

        start_raw, end_raw, lunch_minutes, site = parsed_turno

        in_dt = parse_iso(datetime.strptime(start_raw.replace("am"," am").replace("pm"," pm"), "%I:%M %p").replace(year=datetime.now().year, month=datetime.now().month, day=datetime.now().day).isoformat()) if ":" in start_raw else None

        out_dt = parse_iso(datetime.strptime(end_raw.replace("am"," am").replace("pm"," pm"), "%I:%M %p").replace(year=datetime.now().year, month=datetime.now().month, day=datetime.now().day).isoformat()) if ":" in end_raw else None

        if not in_dt or not out_dt:
            return "Formato inválido. Usa: Turno 8:00 am 5:30 pm 30 lonche Rancho"

        create_shift(phone, name, loc_description=site)

        total_minutes = int((out_dt - in_dt).total_seconds() // 60) - lunch_minutes
        total_minutes = max(total_minutes, 0)

        conn = get_conn()
        shift = conn.execute("SELECT id FROM shifts WHERE phone = ? AND status='open' ORDER BY id DESC LIMIT 1", (phone,)).fetchone()

        conn.execute("""
            UPDATE shifts SET
            clock_out_utc = ?,
            lunch_minutes = ?,
            location_description_out = ?,
            total_work_minutes = ?,
            status = 'closed'
            WHERE id = ?
        """, (
            out_dt.isoformat(),
            lunch_minutes,
            site,
            total_minutes,
            shift["id"],
        ))

        conn.commit()
        conn.close()

        return (
            f"✅ Turno guardado en un mensaje.
"
            f"Entrada: *{in_dt.strftime('%I:%M %p')}*
"
            f"Salida: *{out_dt.strftime('%I:%M %p')}*
"
            f"Lonche: {lunch_minutes} min
"
            f"Lugar: {site}
"
            f"Total trabajado: *{fmt_minutes(total_minutes)}*"
        )

    lower = text.lower().strip()
    name = get_employee_name(phone)
    lower = text.lower().strip()
    name = get_employee_name(phone)

    if lower in {"help", "menu", "ayuda"}:
        return help_text()

    if lower == "start":
        set_state(phone, "awaiting_name")
        return "Send your full name to register this WhatsApp number."

    if lower == "status":
        shift = get_open_shift(phone)
        if not shift:
            return "You do not have an open shift. Send *in* to clock in."
        return (
            f"Open shift found for *{shift['employee_name'] or 'Employee'}*.\n"
            f"Date: {shift['date_local']}\n"
            f"Clock in: {shift['clock_in_utc']} UTC\n"
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

    if lower == "out":
        if not get_open_shift(phone):
            return "No open shift found. Send *in* first."
        set_state(phone, "awaiting_out_lunch")
        return "Send lunch minutes only. Example: *30*"

    if lower.startswith("lunch "):
        try:
            minutes = int(lower.split(" ", 1)[1].strip())
        except ValueError:
            return "Use: *lunch 30*"
        shift = get_open_shift(phone)
        if not shift:
            return "No open shift found. Send *in* first."
        conn = get_conn()
        conn.execute(
            "UPDATE shifts SET lunch_minutes = ? WHERE id = ?",
            (minutes, shift["id"]),
        )
        conn.commit()
        conn.close()
        return f"Lunch updated to *{minutes}* minutes for your current shift."

    if lower.startswith("site "):
        site_text = text[5:].strip()
        shift = get_open_shift(phone)
        if not shift:
            return "No open shift found. Send *in* first."
        conn = get_conn()
        conn.execute(
            "UPDATE shifts SET location_description_in = ? WHERE id = ?",
            (site_text, shift["id"]),
        )
        conn.commit()
        conn.close()
        return f"Site saved: *{site_text}*"

    if lower.startswith("note "):
        note_text = text[5:].strip()
        shift = get_open_shift(phone)
        if not shift:
            return "No open shift found. Send *in* first."
        conn = get_conn()
        conn.execute("UPDATE shifts SET notes = ? WHERE id = ?", (note_text, shift["id"]))
        conn.commit()
        conn.close()
        return "Note saved to your current shift."

    if lower == "report":
        if phone not in ADMIN_NUMBERS:
            return "This command is only available for admin numbers."
        rows = latest_closed_shifts(limit=10)
        if not rows:
            return "No closed shifts yet."
        lines = ["*Latest closed shifts*\n"]
        for row in rows:
            lines.append(
                f"• {row['employee_name'] or row['phone']} | {row['date_local']} | {fmt_minutes(row['total_work_minutes'] or 0)}"
            )
        return "\n".join(lines)

    return (
        "I did not understand that.\n\n"
        f"{help_text()}"
    )


# -----------------------------
# Web routes
# -----------------------------
@app.route("/")
def home():
    return jsonify({
        "app": APP_NAME,
        "ok": True,
        "message": "Use /health, /dashboard?token=YOUR_TOKEN or /export.xlsx?token=YOUR_TOKEN",
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"ok": True, "app": APP_NAME, "timezone": TIMEZONE_LABEL})


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
            body {
                margin: 0;
                font-family: Arial, sans-serif;
                background: #0b1020;
                color: #eef2ff;
            }
            .wrap {
                max-width: 1400px;
                margin: 0 auto;
                padding: 24px;
            }
            .title {
                font-size: 32px;
                font-weight: 700;
                margin-bottom: 6px;
            }
            .sub {
                color: #b7c0d8;
                margin-bottom: 24px;
            }
            .cards {
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 16px;
                margin-bottom: 24px;
            }
            .card {
                background: #121933;
                border: 1px solid #283255;
                border-radius: 16px;
                padding: 18px;
            }
            .label {
                color: #9fb0db;
                font-size: 13px;
                margin-bottom: 8px;
                text-transform: uppercase;
                letter-spacing: 0.05em;
            }
            .value {
                font-size: 28px;
                font-weight: 700;
            }
            .panel {
                background: #121933;
                border: 1px solid #283255;
                border-radius: 18px;
                padding: 18px;
                margin-bottom: 20px;
            }
            form {
                display: grid;
                grid-template-columns: 1.2fr 1fr 1fr auto auto;
                gap: 12px;
                align-items: end;
            }
            label {
                display: block;
                font-size: 13px;
                color: #9fb0db;
                margin-bottom: 6px;
            }
            input {
                width: 100%;
                padding: 12px 14px;
                border-radius: 12px;
                border: 1px solid #33416f;
                background: #0f1530;
                color: #eef2ff;
            }
            .btn {
                display: inline-block;
                padding: 12px 16px;
                border-radius: 12px;
                border: 0;
                text-decoration: none;
                font-weight: 700;
                cursor: pointer;
            }
            .btn-primary {
                background: #6d7cff;
                color: white;
            }
            .btn-secondary {
                background: #1c2547;
                color: #eef2ff;
                border: 1px solid #33416f;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                font-size: 14px;
            }
            th, td {
                padding: 12px;
                border-bottom: 1px solid #223055;
                text-align: left;
                vertical-align: top;
            }
            th {
                color: #9fb0db;
                font-size: 12px;
                text-transform: uppercase;
                letter-spacing: 0.04em;
            }
            .pill {
                display: inline-block;
                padding: 6px 10px;
                border-radius: 999px;
                font-size: 12px;
                font-weight: 700;
            }
            .open { background: #583d00; color: #ffd977; }
            .closed { background: #0d4b32; color: #8df0bf; }
            .tools {
                display: flex;
                gap: 12px;
                margin-top: 14px;
                flex-wrap: wrap;
            }
            @media (max-width: 980px) {
                .cards { grid-template-columns: 1fr 1fr; }
                form { grid-template-columns: 1fr; }
            }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="title">WhatsApp Clock Dashboard</div>
            <div class="sub">See every shift, filter by employee or date, and export to Excel.</div>

            <div class="cards">
                <div class="card">
                    <div class="label">Employees</div>
                    <div class="value">{{ summary.unique_employees }}</div>
                </div>
                <div class="card">
                    <div class="label">Closed Shifts</div>
                    <div class="value">{{ summary.closed_shifts }}</div>
                </div>
                <div class="card">
                    <div class="label">Open Shifts</div>
                    <div class="value">{{ summary.open_shifts }}</div>
                </div>
                <div class="card">
                    <div class="label">Worked Hours</div>
                    <div class="value">{{ worked_hours }}</div>
                </div>
            </div>

            <div class="panel">
                <form method="get" action="/dashboard">
                    <input type="hidden" name="token" value="{{ token }}">
                    <div>
                        <label>Employee or phone</label>
                        <input name="employee" value="{{ employee }}" placeholder="Daniel or +1555...">
                    </div>
                    <div>
                        <label>From date</label>
                        <input type="date" name="date_from" value="{{ date_from }}">
                    </div>
                    <div>
                        <label>To date</label>
                        <input type="date" name="date_to" value="{{ date_to }}">
                    </div>
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
                            <th>Employee</th>
                            <th>Date</th>
                            <th>In</th>
                            <th>Out</th>
                            <th>Lunch</th>
                            <th>Total</th>
                            <th>Status</th>
                            <th>Sites</th>
                            <th>GPS</th>
                            <th>Notes</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in rows %}
                        <tr>
                            <td>
                                <strong>{{ row['employee_name'] or '-' }}</strong><br>
                                <span>{{ row['phone'] }}</span>
                            </td>
                            <td>{{ row['date_local'] or '-' }}</td>
                            <td>{{ fmt_dt(row['clock_in_utc']) }}</td>
                            <td>{{ fmt_dt(row['clock_out_utc']) }}</td>
                            <td>{{ row['lunch_minutes'] or 0 }} min</td>
                            <td>{{ fmt_minutes(row['total_work_minutes'] or 0) }}</td>
                            <td>
                                <span class="pill {{ 'open' if row['status'] == 'open' else 'closed' }}">{{ row['status'] }}</span>
                            </td>
                            <td>
                                <strong>In:</strong> {{ row['location_description_in'] or '-' }}<br>
                                <strong>Out:</strong> {{ row['location_description_out'] or '-' }}
                            </td>
                            <td>
                                <strong>In:</strong> {{ row['in_lat'] or '-' }}, {{ row['in_lng'] or '-' }}<br>
                                <strong>Out:</strong> {{ row['out_lat'] or '-' }}, {{ row['out_lng'] or '-' }}
                            </td>
                            <td>{{ row['notes'] or '-' }}</td>
                        </tr>
                        {% else %}
                        <tr>
                            <td colspan="10">No shifts found for the current filters.</td>
                        </tr>
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
