import os
import json
import datetime
import random
from flask import Flask, request, jsonify
from flask_cors import CORS
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv

# Always load .env from the same folder as app.py (backend/)
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(_BASE_DIR, ".env"))

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": "*"}})

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "").strip()
CREDS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json").strip()
SHEET1 = os.getenv("SHEET1_NAME", "Sheet1").strip() or "Sheet1"
SHEET2 = os.getenv("SHEET2_NAME", "Sheet2").strip() or "Sheet2"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ── Sheet2 column order (0-indexed)
COL = {
    "Date": 0,
    "Tastype": 1,
    "Business ID": 2,
    "TAT": 3,
    "Task Describtion": 4,
    "Employee Name": 5,
    "Collegaue": 6,
    "Status": 7,
    "ChnageOnStatus": 8,
    "Total DaysRequired": 9,
    "Total Days taken": 10,
    "Task Delivery Status": 11,
    "ID": 12,
}
NUM_COLS = 13


def get_service():
    """Create and return Google Sheets API service.

    Prefers JSON from environment variable GOOGLE_SERVICE_ACCOUNT_JSON
    (for cloud hosting like Render), and falls back to a local file
    specified by GOOGLE_CREDENTIALS_FILE / CREDS_FILE.
    """
    json_env = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

    if json_env:
        info = json.loads(json_env)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds_path = os.path.join(_BASE_DIR, CREDS_FILE)
        creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)

    service = build("sheets", "v4", credentials=creds)
    return service.spreadsheets()


def ensure_sheet2_with_header():
    """
    Ensure Sheet2 exists in the spreadsheet and has the correct header row.
    Returns the spreadsheets() service handle.
    """
    sheets = get_service()

    # Ensure the sheet/tab itself exists
    try:
        meta = sheets.get(spreadsheetId=SPREADSHEET_ID).execute()
    except HttpError as e:
        raise

    existing_titles = [
        s.get("properties", {}).get("title", "")
        for s in meta.get("sheets", [])
    ]

    if SHEET2 not in existing_titles:
        body = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": SHEET2,
                        }
                    }
                }
            ]
        }
        sheets.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

    # Ensure header row exists
    header_range = f"{SHEET2}!A1:M1"
    existing_header = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=header_range,
    ).execute()

    if not existing_header.get("values"):
        header = [
            "Date",
            "Tastype",
            "Business ID",
            "TAT",
            "Task Describtion",
            "Employee Name",
            "Collegaue",
            "Status",
            "ChnageOnStatus",
            "Total DaysRequired",
            "Total Days taken",
            "Task Delivery Status",
            "ID",
        ]
        sheets.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=header_range,
            valueInputOption="RAW",
            body={"values": [header]},
        ).execute()

    return sheets


def safe_get(row, idx, default=""):
    """Safely get a cell value from a row list."""
    try:
        val = row[idx]
        return val if val is not None else default
    except IndexError:
        return default


def parse_date(date_str):
    """Parse common date formats and return a date object."""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(date_str.strip(), fmt).date()
        except (ValueError, AttributeError):
            continue
    return None


def generate_task_id(client_name, worker_name, today):
    """
    Generate Task ID: CLIENTCODE_RANDOMNUM-WorkerName-YYYYMMDD
    Example: ASHRA_75076-Abhishek-20260223
    """
    client_code = client_name.replace(" ", "")[:5].upper()
    random_num = random.randint(10000, 99999)
    date_str = today.strftime("%Y%m%d")
    worker_clean = worker_name.strip().replace(" ", "")
    return f"{client_code}_{random_num}-{worker_clean}-{date_str}"


def get_delivery_status(tat_date, completion_date):
    """Return Task Delivery Status string based on dates."""
    days_late = (completion_date - tat_date).days
    if days_late <= 0:
        return "On Time"
    elif days_late == 1:
        return "Late Submission"
    else:
        return "Late Delivery"


# ─────────────────────────────────────────────
#  GET /api/sheet1
#  Returns workers (col1), clients (col2), task types (col3)
# ─────────────────────────────────────────────
@app.route("/api/sheet1", methods=["GET"])
def get_sheet1():
    try:
        sheets = get_service()
        result = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET1}!A:C"
        ).execute()
        rows = result.get("values", [])

        # Skip header row (assumed first row) so labels like
        # "Worker Names", "Client Names", "Task Types" do not
        # appear in dropdowns.
        data_rows = rows[1:] if rows else []

        workers = []
        clients = []
        task_types = []

        for row in data_rows:
            if len(row) >= 1 and row[0].strip():
                workers.append(row[0].strip())
            if len(row) >= 2 and row[1].strip():
                clients.append(row[1].strip())
            if len(row) >= 3 and row[2].strip():
                task_types.append(row[2].strip())

        return jsonify({
            "workers": list(dict.fromkeys(workers)),      # deduplicate, preserve order
            "clients": list(dict.fromkeys(clients)),
            "taskTypes": list(dict.fromkeys(task_types))
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────
#  GET /api/tasks
#  Returns ALL rows from Sheet2
# ─────────────────────────────────────────────
@app.route("/api/tasks", methods=["GET"])
def get_tasks():
    try:
        sheets = ensure_sheet2_with_header()
        result = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET2}!A:M"
        ).execute()
        rows = result.get("values", [])

        if not rows:
            return jsonify({"tasks": []})

        # First row is headers — skip it
        headers = rows[0] if rows else []
        tasks = []
        for i, row in enumerate(rows[1:], start=2):  # start=2 for sheet row number
            task = {
                "rowIndex": i,
                "Date": safe_get(row, COL["Date"]),
                "Tastype": safe_get(row, COL["Tastype"]),
                "Business ID": safe_get(row, COL["Business ID"]),
                "TAT": safe_get(row, COL["TAT"]),
                "Task Describtion": safe_get(row, COL["Task Describtion"]),
                "Employee Name": safe_get(row, COL["Employee Name"]),
                "Collegaue": safe_get(row, COL["Collegaue"]),
                "Status": safe_get(row, COL["Status"]),
                "ChnageOnStatus": safe_get(row, COL["ChnageOnStatus"]),
                "Total DaysRequired": safe_get(row, COL["Total DaysRequired"]),
                "Total Days taken": safe_get(row, COL["Total Days taken"]),
                "Task Delivery Status": safe_get(row, COL["Task Delivery Status"]),
                "ID": safe_get(row, COL["ID"]),
            }
            tasks.append(task)

        return jsonify({"tasks": tasks})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────
#  GET /api/tasks/active
#  Returns only non-Completed, non-Cancelled tasks
# ─────────────────────────────────────────────
@app.route("/api/tasks/active", methods=["GET"])
def get_active_tasks():
    try:
        sheets = ensure_sheet2_with_header()
        result = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET2}!A:M"
        ).execute()
        rows = result.get("values", [])

        if not rows:
            return jsonify({"tasks": []})

        excluded = {"Completed", "Cancelled"}
        tasks = []
        for i, row in enumerate(rows[1:], start=2):
            status = safe_get(row, COL["Status"]).strip()
            if status in excluded:
                continue
            task = {
                "rowIndex": i,
                "Date": safe_get(row, COL["Date"]),
                "Tastype": safe_get(row, COL["Tastype"]),
                "Business ID": safe_get(row, COL["Business ID"]),
                "TAT": safe_get(row, COL["TAT"]),
                "Task Describtion": safe_get(row, COL["Task Describtion"]),
                "Employee Name": safe_get(row, COL["Employee Name"]),
                "Collegaue": safe_get(row, COL["Collegaue"]),
                "Status": safe_get(row, COL["Status"]),
                "ChnageOnStatus": safe_get(row, COL["ChnageOnStatus"]),
                "Total DaysRequired": safe_get(row, COL["Total DaysRequired"]),
                "Total Days taken": safe_get(row, COL["Total Days taken"]),
                "Task Delivery Status": safe_get(row, COL["Task Delivery Status"]),
                "ID": safe_get(row, COL["ID"]),
            }
            tasks.append(task)

        return jsonify({"tasks": tasks})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────
#  POST /api/tasks/create
#  Appends a new task row to Sheet2
# ─────────────────────────────────────────────
@app.route("/api/tasks/create", methods=["POST"])
def create_task():
    try:
        data = request.get_json()

        # Validate required fields
        required = ["taskType", "clientId", "tat", "taskDescription", "workerName"]
        for field in required:
            if not data.get(field, "").strip():
                return jsonify({"error": f"Missing required field: {field}"}), 400

        today = datetime.date.today()
        today_str = today.strftime("%Y-%m-%d")

        task_type = data["taskType"].strip()
        client_id = data["clientId"].strip()
        tat_str = data["tat"].strip()
        description = data["taskDescription"].strip()
        worker_name = data["workerName"].strip()
        colleague = data.get("colleague", "NONE").strip() or "NONE"

        # Calculate Total Days Required
        tat_date = parse_date(tat_str)
        if tat_date:
            days_required = (tat_date - today).days
        else:
            days_required = ""

        # Generate Task ID
        task_id = generate_task_id(client_id, worker_name, today)

        # Build row in exact Sheet2 column order
        row = [""] * NUM_COLS
        row[COL["Date"]] = today_str
        row[COL["Tastype"]] = task_type
        row[COL["Business ID"]] = client_id
        row[COL["TAT"]] = tat_str
        row[COL["Task Describtion"]] = description
        row[COL["Employee Name"]] = worker_name
        row[COL["Collegaue"]] = colleague
        row[COL["Status"]] = "Pending"
        row[COL["ChnageOnStatus"]] = ""
        row[COL["Total DaysRequired"]] = str(days_required) if days_required != "" else ""
        row[COL["Total Days taken"]] = ""
        row[COL["Task Delivery Status"]] = ""
        row[COL["ID"]] = task_id

        sheets = ensure_sheet2_with_header()

        # Append the data row
        sheets.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET2}!A:M",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]}
        ).execute()

        return jsonify({
            "success": True,
            "taskId": task_id,
            "message": f"Task created successfully! ID: {task_id}"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────
#  PUT /api/tasks/update
#  Updates an existing row in Sheet2 by Task ID
# ─────────────────────────────────────────────
@app.route("/api/tasks/update", methods=["PUT"])
def update_task():
    try:
        data = request.get_json()

        task_id = data.get("taskId", "").strip()
        new_status = data.get("newStatus", "").strip()

        if not task_id:
            return jsonify({"error": "taskId is required"}), 400
        if not new_status:
            return jsonify({"error": "newStatus is required"}), 400

        sheets = get_service()

        # Ensure Sheet2 and header exist, then read all rows to find the target
        sheets = ensure_sheet2_with_header()
        result = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET2}!A:M"
        ).execute()
        rows = result.get("values", [])

        target_row_index = None  # 1-indexed sheet row
        target_row = None
        for i, row in enumerate(rows):
            if safe_get(row, COL["ID"]).strip() == task_id:
                target_row_index = i + 1  # Sheets is 1-indexed
                target_row = list(row) + [""] * (NUM_COLS - len(row))
                break

        if target_row is None:
            return jsonify({"error": f"Task ID '{task_id}' not found"}), 404

        today = datetime.date.today()
        today_str = today.strftime("%Y-%m-%d")

        # Update status and change date
        target_row[COL["Status"]] = new_status
        target_row[COL["ChnageOnStatus"]] = today_str

        # Reassign worker/colleague if provided
        if data.get("newWorker"):
            target_row[COL["Employee Name"]] = data["newWorker"].strip()
        if data.get("newColleague"):
            target_row[COL["Collegaue"]] = data["newColleague"].strip()

        # If Completed or Cancelled — calculate total days taken and delivery status
        if new_status in ("Completed", "Cancelled"):
            assigned_date_str = safe_get(target_row, COL["Date"])
            tat_str = safe_get(target_row, COL["TAT"])

            assigned_date = parse_date(assigned_date_str)
            tat_date = parse_date(tat_str)

            if assigned_date:
                days_taken = (today - assigned_date).days
                target_row[COL["Total Days taken"]] = str(days_taken)

            if tat_date:
                delivery_status = get_delivery_status(tat_date, today)
                target_row[COL["Task Delivery Status"]] = delivery_status

        # Pad row to ensure it covers all columns
        while len(target_row) < NUM_COLS:
            target_row.append("")

        # Write back the updated row
        range_notation = f"{SHEET2}!A{target_row_index}:M{target_row_index}"
        sheets.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_notation,
            valueInputOption="RAW",
            body={"values": [target_row[:NUM_COLS]]}
        ).execute()

        return jsonify({
            "success": True,
            "message": f"Task '{task_id}' updated to '{new_status}' successfully."
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────
#  Root route — status page (so browser doesn't show 404)
# ─────────────────────────────────────────────
@app.route("/", methods=["GET"])
def root():
    sheet_ok = bool(SPREADSHEET_ID and SPREADSHEET_ID != "paste_your_spreadsheet_id_here")
    return f"""
    <html><head><title>Reputes Work Tracker API</title>
    <style>body{{font-family:Arial,sans-serif;background:#0d47d9;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;}}
    .box{{background:rgba(255,255,255,0.12);border-radius:16px;padding:40px 50px;text-align:center;box-shadow:0 4px 32px rgba(0,0,0,0.3);}}
    h1{{font-size:2em;margin-bottom:8px;letter-spacing:2px;}}p{{opacity:.85;margin:6px 0;}}
    .ok{{color:#90ee90;font-weight:bold;}}.warn{{color:#ffcc00;font-weight:bold;}}
    a{{color:#90caf9;}}ul{{text-align:left;margin-top:16px;line-height:2;}}</style></head>
    <body><div class='box'>
    <h1>🚀 REPUTES WORK TRACKER</h1>
    <p>Flask API Server is <span class='ok'>RUNNING ✅</span></p>
    <p>Spreadsheet ID: <span class='{'ok' if sheet_ok else 'warn'}'>{'Set ✅' if sheet_ok else 'NOT SET ⚠️ — edit backend/.env'}</span></p>
    <p style='margin-top:20px;font-size:.9em;opacity:.7'>Available API Endpoints:</p>
    <ul>
    <li><a href='/api/health'>/api/health</a> — Health check</li>
    <li><a href='/api/sheet1'>/api/sheet1</a> — Master data (workers, clients, task types)</li>
    <li><a href='/api/tasks'>/api/tasks</a> — All tasks</li>
    <li><a href='/api/tasks/active'>/api/tasks/active</a> — Active tasks only</li>
    </ul>
    </div></body></html>
    """, 200


# ─────────────────────────────────────────────
#  Health check
# ─────────────────────────────────────────────
@app.route("/api/health", methods=["GET"])
def health():
    sheet_ok = bool(SPREADSHEET_ID and SPREADSHEET_ID != "paste_your_spreadsheet_id_here")
    return jsonify({
        "status": "ok",
        "message": "Reputes Work Tracker API is running",
        "spreadsheet_configured": sheet_ok,
        "sheet1": SHEET1,
        "sheet2": SHEET2
    })


if __name__ == "__main__":
    port = int(os.getenv("FLASK_PORT", 5000))
    print(f"\n🚀 Reputes Work Tracker API starting on http://0.0.0.0:{port}")
    print(f"📊 Spreadsheet ID: {SPREADSHEET_ID or 'NOT SET — edit backend/.env'}")
    print(f"🔑 Credentials file: {CREDS_FILE}\n")
    app.run(host="0.0.0.0", port=port, debug=True)
