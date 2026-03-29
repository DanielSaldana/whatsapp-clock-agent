import os
import sqlite3
from datetime import datetime, timezone
from flask import Flask, request, jsonify
from twilio.twiml.messaging_response import MessagingResponse

APP_NAME = "WhatsApp Clock Agent"
DB_PATH = os.getenv("DB_PATH", "timeclock.db")
ADMIN_NUMBERS = {
    n.strip() for n in os.getenv("ADMIN_NUMBERS", "").split(",") if n.strip()
}
TIMEZONE_LABEL = os.getenv("TIMEZONE_LABEL", "America/Denver")

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


# Keeping local date simple for MVP.
# In production, convert to the user's/business timezone using zoneinfo.
def local_date_string() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def normalize_text(value: str) -> str:
    return (value or "").strip()


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


def fmt_minutes(total_minutes: int) -> str:
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}h {minutes}m"


def latest_closed_shifts(limit: int = 20):
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT employee_name, phone, date_local, lunch_minutes, total_work_minutes,
               location_description_in, location_description_out,
               clock_in_utc, clock_out_utc
        FROM shifts
        WHERE status = 'closed'
        ORDER BY id DESC
        LIMIT ?
        """,
        (limit,),
    ).fetchall()
    conn.close()
    return rows


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
            "UPDATE shifts SET location_description_in = COALESCE(location_description_in, ?) WHERE id = ?",
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


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"ok": True, "app": APP_NAME, "timezone": TIMEZONE_LABEL})


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
