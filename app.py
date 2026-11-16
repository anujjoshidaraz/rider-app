"""
Barcode Tracker — lightweight local web app
Scans barcodes via browser camera, records delivery issue reasons,
and stores everything in DingTalk Sheet (dingsheet).
Uses dws CLI (already authenticated) — no API keys needed.
Run: python app.py  →  opens http://localhost:57801 in your browser
"""

import json
import os
import subprocess
import sys
import time
import threading
import webbrowser
from datetime import datetime, timezone, timedelta

from flask import Flask, request, jsonify, Response, session, redirect, url_for, send_from_directory

# Import WhatsApp automation module (same directory)
try:
    from whatsapp_sender import send_message_to_group, is_logged_in, setup_login
    WHATSAPP_AVAILABLE = True
    print("WhatsApp module loaded successfully")
except Exception as e:
    WHATSAPP_AVAILABLE = False
    print(f"Warning: WhatsApp module failed to load: {type(e).__name__}: {e}")

# ─── Configuration ────────────────────────────────────────────────────────────
NODE_ID = "EpGBa2Lm8azLbBLwTzjp705DWgN7R35y"
SHEET_ID = "st-52d05d2f-402912"     # Scans subsheet (driver scan data with reasons)
RIDERS_SHEET_ID = "st-be62cc2f-247786"  # Riders directory sheet (Riders tab)
COMMENTS_SHEET_ID = "st-700fb79b-272759"  # Comments sheet (from dashboard)
SHEET_URL = "https://alidocs.dingtalk.com/i/nodes/EpGBa2Lm8azLbBLwTzjp705DWgN7R35y"

# Row counter: headers are row 1, data starts at row 2
_next_row = 2

PORT = int(os.environ.get('PORT', 57801))
ADMIN_PASSWORD = "daraz2026"  # Change this to your preferred admin password

# ─── Locate dws-core executable ───────────────────────────────────────────────
def _find_dws_core():
    """Find the dws-core executable — works on Windows, Linux, and macOS."""
    import platform
    home = os.path.expanduser("~")
    system = platform.system().lower()
    
    candidates = []
    
    if system == 'windows':
        candidates = [
            os.path.join(home, ".qoderwork", "bin", "ext", "dws-core-windows-amd64.exe"),
        ]
        if getattr(sys, "frozen", False):
            candidates.append(os.path.join(os.path.dirname(sys.executable), "dws-core-windows-amd64.exe"))
    elif system == 'linux':
        candidates = [
            os.path.join(home, ".qoderwork", "bin", "dws"),
            '/usr/local/bin/dws',
            os.path.join(home, "dws"),
        ]
    else:  # macOS
        candidates = [
            os.path.join(home, ".qoderwork", "bin", "dws"),
            '/usr/local/bin/dws',
        ]
    
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None

DWS_CORE = _find_dws_core()


def _run_dws(args, timeout=60):
    """Run a dws-core command and return parsed JSON."""
    if not DWS_CORE:
        raise Exception("dws-core not found. Make sure QoderWork is installed.")
    cmd = [DWS_CORE] + args
    result = subprocess.run(
        cmd,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )
    stdout = result.stdout or ""
    stderr = result.stderr or ""
    if result.returncode != 0:
        err_msg = (stderr.strip() or stdout.strip() or f"exit code {result.returncode}")
        raise Exception(f"dws error: {err_msg}")
    try:
        return json.loads(stdout)
    except json.JSONDecodeError:
        return {"raw": stdout}


# ─── Fetch riders from dingsheet ──────────────────────────────────────────────

# ─── Fetch riders from dingsheet ──────────────────────────────────────────────

def fetch_riders():
    """Read the Riders sheet tab and return list of {hub, rider, active}."""
    result = _run_dws([
        "sheet", "range", "read",
        "--node", NODE_ID,
        "--sheet-id", RIDERS_SHEET_ID,
        "--format", "json",
    ])
    cells = result.get("cells", [])
    if len(cells) < 2:
        return []
    # Skip header row (index 0), parse data rows
    riders = []
    for row in cells[1:]:
        if len(row) >= 3:
            hub = row[0].get("value", "").strip()
            rider = row[1].get("value", "").strip()
            active = row[2].get("value", "").strip().lower()
            if hub and rider and active in ("yes", "true", "1", "y"):
                riders.append({"hub": hub, "rider": rider})
    return riders


# ─── Sync row counter with actual sheet data ──────────────────────────────────

def sync_row_counter():
    """Query the sheet to find the actual last row with data, set _next_row accordingly."""
    global _next_row
    try:
        result = _run_dws([
            "sheet", "info",
            "--node", NODE_ID,
            "--sheet-id", SHEET_ID,
            "--format", "json",
        ])
        non_empty = result.get("nonEmptyRange", {})
        last_row = non_empty.get("lastRow", 1)
        _next_row = last_row + 1
        print(f"  Row counter synced: next row = {_next_row}")
    except Exception as e:
        print(f"  Warning: could not sync row counter ({e}), defaulting to row 2")
        _next_row = 2


# ─── Record creation (dingsheet) ──────────────────────────────────────────────

def create_record(rider_name, hub, tracking, scanned_at_ms, status="Pending", notes="", reason_category="", sub_reason=""):
    """Append a row to the DingTalk sheet with rider/hub + scan data + reason."""
    global _next_row

    # Re-sync row counter before each write to stay accurate
    sync_row_counter()

    # Convert epoch ms to readable timestamp (NPT = UTC+5:45)
    npt = timezone(timedelta(hours=5, minutes=45))
    ts = datetime.fromtimestamp(scanned_at_ms / 1000, tz=npt)
    scanned_at_str = ts.strftime("%Y-%m-%d %H:%M:%S")

    # Append text data: columns A-H (Rider Name, Hub, Tracking, Scanned At, Status, Reason Category, Sub-Reason, Notes)
    values = json.dumps([[rider_name, hub, tracking, scanned_at_str, status, reason_category, sub_reason, notes or ""]])
    _run_dws([
        "sheet", "append",
        "--node", NODE_ID,
        "--sheet-id", SHEET_ID,
        "--values", values,
        "--format", "json",
    ])

    current_row = _next_row
    _next_row += 1

    return {"success": True, "row": current_row}


# ─── Admin: Read all records from dingsheet ──────────────────────────────────

_CACHE_TTL = 8  # seconds
_records_cache = {"data": None, "ts": 0}
_comments_cache = {"data": None, "ts": 0}

def fetch_all_records():
    """Read all scan records from the main dingsheet (cached)."""
    now = time.time()
    if _records_cache["data"] is not None and (now - _records_cache["ts"]) < _CACHE_TTL:
        return _records_cache["data"]
    result = _run_dws([
        "sheet", "range", "read",
        "--node", NODE_ID,
        "--sheet-id", SHEET_ID,
        "--format", "json",
    ])
    cells = result.get("cells", [])
    if len(cells) < 2:
        _records_cache["data"] = []
        _records_cache["ts"] = now
        return []
    # Headers: Rider Name(A), Hub(B), Tracking(C), Scanned At(D), Status(E), Reason Category(F), Sub-Reason(G), Notes(H), Comments(I)
    records = []
    for idx, row in enumerate(cells[1:], start=2):  # row 2 is first data row
        if len(row) < 4:
            continue
        rider = row[0].get("value", "").strip() if len(row) > 0 else ""
        hub = row[1].get("value", "").strip() if len(row) > 1 else ""
        tracking = row[2].get("value", "").strip() if len(row) > 2 else ""
        scanned_at = row[3].get("value", "").strip() if len(row) > 3 else ""
        status = row[4].get("value", "").strip() if len(row) > 4 else "Pending"
        reason_category = row[5].get("value", "").strip() if len(row) > 5 else ""
        sub_reason = row[6].get("value", "").strip() if len(row) > 6 else ""
        notes = row[7].get("value", "").strip() if len(row) > 7 else ""
        comments = row[8].get("value", "").strip() if len(row) > 8 else ""
        if tracking:  # Skip empty rows
            records.append({
                "row": idx,
                "rider": rider,
                "hub": hub,
                "tracking": tracking,
                "scanned_at": scanned_at,
                "status": status,
                "reason_category": reason_category,
                "sub_reason": sub_reason,
                "notes": notes,
                "comments": comments,
            })
    _records_cache["data"] = records
    _records_cache["ts"] = now
    return records

def fetch_comments_from_sheet():
    """Fetch all comments from the Comments sheet (cached).
    Returns dict: {tracking: combined_comment_string}
    """
    now = time.time()
    if _comments_cache["data"] is not None and (now - _comments_cache["ts"]) < _CACHE_TTL:
        return _comments_cache["data"]
    result = _run_dws([
        "sheet", "range", "read",
        "--node", NODE_ID,
        "--sheet-id", COMMENTS_SHEET_ID,
        "--range", "A2:D5000",
        "--format", "json",
    ])
    if "error" in result:
        _comments_cache["data"] = {}
        _comments_cache["ts"] = now
        return {}
    cells = result.get("cells", [])
    comments_map = {}
    for row in cells:
        if not row or not row[0].get("value", "").strip():
            continue
        tracking = row[0].get("value", "").strip()
        comment = row[1].get("value", "").strip() if len(row) > 1 else ""
        author = row[2].get("value", "").strip() if len(row) > 2 else ""
        timestamp = row[3].get("value", "").strip() if len(row) > 3 else ""
        if tracking and comment:
            entry = f"[{timestamp}] {comment}"
            if author:
                entry += f" ({author})"
            if tracking in comments_map:
                comments_map[tracking] += "\n" + entry
            else:
                comments_map[tracking] = entry
    _comments_cache["data"] = comments_map
    _comments_cache["ts"] = now
    return comments_map


def update_record_status(row_num, new_status):
    """Update the status cell (column E) for a given row."""
    values = json.dumps([[{"type": "text", "text": new_status}]])
    _run_dws([
        "sheet", "range", "update",
        "--node", NODE_ID,
        "--sheet-id", SHEET_ID,
        "--range", f"E{row_num}",
        "--values", values,
        "--format", "json",
    ])
    return {"success": True, "row": row_num, "status": new_status}


# ─── Flask App ────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = "barcode-tracker-secret-key-2026"  # Change this in production

# ─── Authentication ──────────────────────────────────────────────────────────
AUTH_NODE_ID = "G1DKw2zgV2RNb5N7Td5z2dpPVB5r9YAn"  # Separate node for rider credentials
AUTH_SHEET_ID = "st-301c4f0c-121671"  # WhatsApp Admin sheet with credentials

def get_rider_credentials():
    """Read rider credentials from DingTalk Sheet.
    Sheet structure: A=Phone, B=Name, C=Password
    Returns dict: {phone_number: {"password": password, "name": name}}
    """
    result = _run_dws(["sheet", "range", "read", "--node", AUTH_NODE_ID,
                      "--sheet-id", AUTH_SHEET_ID, "--range", "A2:C100"])
    if "error" in result:
        return {}
    cells = result.get("cells", [])
    credentials = {}
    for row in cells:
        if not row or not row[0].get("value", "").strip():
            continue
        phone = row[0].get("value", "").strip()
        rider_name = row[1].get("value", "").strip() if len(row) > 1 else ""
        password = row[2].get("value", "").strip() if len(row) > 2 else ""
        if phone and password:
            credentials[phone] = {"password": password, "name": rider_name}
    return credentials

def register_rider(phone, name, password):
    """Register a new rider in the auth sheet.
    Sheet structure: A=Phone, B=Name, C=Password
    """
    # Find next empty row
    result = _run_dws(["sheet", "range", "read", "--node", AUTH_NODE_ID,
                      "--sheet-id", AUTH_SHEET_ID, "--range", "A:A"])
    next_row = 2
    if "error" not in result:
        existing = result.get("cells", [])
        for i, row in enumerate(existing):
            if not row or not row[0].get("value", "").strip():
                next_row = i + 2
                break
        else:
            next_row = len(existing) + 2
    
    values = [[
        {"type": "text", "text": phone},
        {"type": "text", "text": name},
        {"type": "text", "text": password}
    ]]
    write_result = _run_dws(["sheet", "range", "update", "--node", AUTH_NODE_ID,
                           "--sheet-id", AUTH_SHEET_ID, "--range", f"A{next_row}:C{next_row}",
                           "--values", json.dumps(values), "--timeout", "60"])
    if "error" in write_result:
        return write_result
    return {"ok": True, "row": next_row}

def check_phone_exists(phone):
    """Check if a phone number already exists in the auth sheet."""
    credentials = get_rider_credentials()
    return phone in credentials

def login_required(f):
    """Decorator to protect routes that require authentication."""
    from functools import wraps
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("authenticated"):
            # For API routes, return 401 JSON
            if request.path.startswith("/api/"):
                return jsonify({"error": "Unauthorized", "login_required": True}), 401
            # For page routes, redirect to login
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated_function


@app.route("/logo.png")
def logo():
    return send_from_directory(os.path.dirname(os.path.abspath(__file__)), "daraz_logo.png")


@app.route("/")
def index():
    # Inject timestamp so the browser always sees the page as new content
    html_with_ts = HTML_PAGE.replace("</title>", f"</title>\n<!-- v{int(time.time())} -->")
    resp = Response(html_with_ts, content_type="text/html")
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    resp.headers["ETag"] = str(time.time())
    return resp


@app.route("/api/riders", methods=["GET"])
def api_riders():
    """Return the list of active riders grouped by hub from the dingsheet."""
    try:
        riders = fetch_riders()
        # Group by hub
        hubs = {}
        for r in riders:
            hubs.setdefault(r["hub"], []).append(r["rider"])
        return jsonify({"ok": True, "hubs": hubs})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/scan", methods=["POST"])
def api_scan():
    data = request.json
    rider_name = (data.get("rider_name") or "").strip()
    hub = (data.get("hub") or "").strip()
    tracking = (data.get("tracking") or "").strip()
    if not tracking:
        return jsonify({"ok": False, "error": "No tracking number"}), 400
    if not rider_name or not hub:
        return jsonify({"ok": False, "error": "Please login first (select hub and rider)"}), 400
    scanned_at = data.get("scanned_at", int(time.time() * 1000))
    status = data.get("status", "Pending")
    notes = data.get("notes", "")
    reason_category = (data.get("reason_category") or "").strip()
    sub_reason = (data.get("sub_reason") or "").strip()
    try:
        result = create_record(rider_name, hub, tracking, scanned_at, status, notes, reason_category, sub_reason)
        return jsonify({"ok": True, "result": str(result)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/check-rescan", methods=["POST"])
def api_check_rescan():
    """Check if a tracking number is a rescan (exists with status=Completed).
    Used to distinguish handoff scans from fresh scans."""
    data = request.json
    tracking = (data.get("tracking") or "").strip()
    if not tracking:
        return jsonify({"ok": False, "error": "tracking required"}), 400
    try:
        all_records = fetch_all_records()
        # Check if any row for this tracking has status=Completed
        is_rescan = any(
            r["tracking"] == tracking and
            str(r.get("status") or "").lower() == "completed"
            for r in all_records
        )
        return jsonify({"ok": True, "is_rescan": is_rescan})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/my-records", methods=["GET"])
def api_my_records():
    """Return scan records for a specific rider, including comments from the Comments sheet."""
    rider = (request.args.get("rider") or "").strip()
    if not rider:
        return jsonify({"ok": False, "error": "rider query param required"}), 400
    try:
        all_records = fetch_all_records()
        # Fetch comments from the Comments sheet
        comments_map = fetch_comments_from_sheet()
        # Filter to only this rider's records
        my_records = [r for r in all_records if r["rider"].lower() == rider.lower()]
        # A rescan creates a NEW sheet row for the same tracking number, so the same
        # parcel shows up multiple times. Collapse to the newest row per tracking, but
        # merge every comment line ever attached to that tracking (column I of all of
        # its rows plus the Comments sheet) so no feedback is lost.
        merged_lines = {}
        for r in my_records:
            tracking = r["tracking"]
            bucket = merged_lines.setdefault(tracking, [])
            lines = []
            if r["comments"]:
                lines.extend(r["comments"].split("\n"))
            sheet_comments = comments_map.get(tracking)
            if sheet_comments:
                lines.extend(sheet_comments.split("\n"))
            for line in lines:
                stripped = line.strip()
                if stripped and stripped not in bucket:
                    bucket.append(stripped)
        newest = {}
        for r in my_records:
            key = r["tracking"]
            if key not in newest or r["row"] > newest[key]["row"]:
                newest[key] = r
        collapsed = []
        for r in newest.values():
            r["comments"] = "\n".join(merged_lines.get(r["tracking"], []))
            collapsed.append(r)
        my_records = collapsed
        # Sort by row (newest first)
        my_records.sort(key=lambda r: r["row"], reverse=True)
        return jsonify({"ok": True, "records": my_records, "count": len(my_records)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/rider-comment", methods=["POST"])
def api_rider_comment():
    """Add a comment from a rider to the Comments sheet.
    Accepts {rider, tracking, comment} JSON body.
    Writes: tracking | comment | rider_name | timestamp
    """
    data = request.json
    rider = (data.get("tracking") or "").strip()
    tracking = (data.get("tracking") or "").strip()
    comment = (data.get("comment") or "").strip()
    if not tracking:
        return jsonify({"ok": False, "error": "tracking required"}), 400
    if not comment:
        return jsonify({"ok": False, "error": "comment required"}), 400
    rider_name = (data.get("rider") or "").strip()
    if not rider_name:
        return jsonify({"ok": False, "error": "rider required"}), 400
    try:
        npt = timezone(timedelta(hours=5, minutes=45))
        timestamp = datetime.now(npt).strftime("%Y-%m-%d %H:%M:%S")
        # Find next empty row in Comments sheet
        read_result = _run_dws([
            "sheet", "range", "read",
            "--node", NODE_ID,
            "--sheet-id", COMMENTS_SHEET_ID,
            "--range", "A1:A1000",
            "--format", "json",
        ])
        next_row = 2
        existing = read_result.get("cells", [])
        for i, row in enumerate(existing):
            if not row or not row[0].get("value", "").strip():
                next_row = i + 1
                break
        else:
            next_row = len(existing) + 2 if existing else 2
        # Write the comment row: A=tracking, B=comment, C=author(rider name), D=timestamp
        values = json.dumps([[
            {"type": "text", "text": tracking},
            {"type": "text", "text": comment},
            {"type": "text", "text": rider_name},
            {"type": "text", "text": timestamp}
        ]])
        range_str = f"A{next_row}:D{next_row}"
        _run_dws([
            "sheet", "range", "update",
            "--node", NODE_ID,
            "--sheet-id", COMMENTS_SHEET_ID,
            "--range", range_str,
            "--values", values,
            "--timeout", "60",
        ])
        # Invalidate caches so next read gets fresh data
        _records_cache["data"] = None
        _records_cache["ts"] = 0
        _comments_cache["data"] = None
        _comments_cache["ts"] = 0
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/end-package", methods=["POST"])
def api_end_package():
    """Mark a package as Completed for a rider. Updates status in Barcode Tracker sheet."""
    data = request.json
    rider = (data.get("rider") or "").strip()
    tracking = (data.get("tracking") or "").strip()
    if not rider or not tracking:
        return jsonify({"ok": False, "error": "Missing rider or tracking"}), 400
    try:
        all_records = fetch_all_records()
        # Find the LATEST record for this rider+tracking.
        # fetch_all_records returns rows in ascending sheet order, and a rescan adds a
        # new row for the same tracking, so we must pick the highest row number.
        target = None
        for r in all_records:
            if r["rider"].lower() == rider.lower() and r["tracking"] == tracking:
                if target is None or r["row"] > target["row"]:
                    target = r
        if not target:
            return jsonify({"ok": False, "error": "Record not found"}), 404
        result = update_record_status(target["row"], "Completed")
        # Invalidate cache
        _records_cache["data"] = None
        _records_cache["ts"] = 0
        return jsonify({"ok": True, "result": result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ─── Admin Routes ─────────────────────────────────────────────────────────────

@app.route("/admin")
def admin_page():
    return Response(ADMIN_HTML, content_type="text/html")


@app.route("/api/admin/login", methods=["POST"])
def admin_login():
    data = request.json
    password = data.get("password", "")
    if password == ADMIN_PASSWORD:
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Wrong password"}), 401


# ── Rider Authentication ───────────────────────────────────────────────────

LOGIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Rider Login - Barcode Tracker</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 20px;
}
.login-container {
  background: white;
  border-radius: 12px;
  padding: 40px;
  width: 100%;
  max-width: 400px;
  box-shadow: 0 10px 40px rgba(0,0,0,0.2);
}
.login-header { text-align: center; margin-bottom: 30px; }
.login-header h1 { font-size: 24px; color: #1a1a1a; margin-bottom: 8px; }
.login-header p { color: #666; font-size: 14px; }
.form-group { margin-bottom: 20px; }
.form-group label { display: block; font-size: 13px; font-weight: 600; color: #333; margin-bottom: 8px; }
.form-group input { width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 14px; }
.form-group input:focus { outline: none; border-color: #667eea; }
.login-btn { width: 100%; padding: 12px; background: #667eea; color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; }
.login-btn:hover { background: #5568d3; }
.login-btn:disabled { background: #ccc; cursor: not-allowed; }
.error-msg { background: #fee; color: #c33; padding: 12px; border-radius: 8px; font-size: 13px; margin-bottom: 20px; display: none; }
.error-msg.show { display: block; }
.signup-link { text-align: center; margin-top: 20px; font-size: 14px; color: #666; }
.signup-link a { color: #667eea; text-decoration: none; font-weight: 600; }
.signup-link a:hover { text-decoration: underline; }
</style>
</head>
<body>
<div class="login-container">
  <div class="login-header"><h1>🔐 Rider Login</h1><p>Barcode Tracker</p></div>
  <div class="error-msg" id="error-msg"></div>
  <form id="login-form">
    <div class="form-group"><label for="phone">Phone Number</label><input type="tel" id="phone" placeholder="e.g., 98XXXXXXXX" required></div>
    <div class="form-group"><label for="pin">PIN</label><input type="password" id="pin" placeholder="Enter PIN" required></div>
    <button type="submit" class="login-btn" id="login-btn">Login</button>
  </form>
  <div class="signup-link">Don't have an account? <a href="/signup">Sign up here</a></div>
</div>
<script>
document.getElementById('login-form').addEventListener('submit', async (e) => {
  e.preventDefault();
  const btn = document.getElementById('login-btn');
  const errorMsg = document.getElementById('error-msg');
  const phone = document.getElementById('phone').value.trim();
  const pin = document.getElementById('pin').value;
  btn.disabled = true; btn.textContent = 'Logging in...'; errorMsg.classList.remove('show');
  try {
    const res = await fetch('/api/rider/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ phone: phone, pin: pin }) });
    const data = await res.json();
    if (data.ok) { window.location.href = '/'; }
    else { errorMsg.textContent = data.error || 'Login failed'; errorMsg.classList.add('show'); btn.disabled = false; btn.textContent = 'Login'; }
  } catch (err) { errorMsg.textContent = 'Network error'; errorMsg.classList.add('show'); btn.disabled = false; btn.textContent = 'Login'; }
});
</script>
</body>
</html>
"""

SIGNUP_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rider Signup - Daraz Barcode Tracker</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: { sans: ['Inter', 'system-ui', 'sans-serif'] },
                    colors: {
                        brand: { 50:'#FFF7ED',100:'#FFEDD5',200:'#FDBA74',300:'#FB923C',400:'#F97316',500:'#F85C2C',600:'#EA580C',700:'#C2410C' },
                    },
                    animation: {
                        'blob': 'blob 8s infinite',
                        'fade-up': 'fadeUp 0.6s cubic-bezier(0.16,1,0.3,1)',
                        'shimmer': 'shimmer 2s linear infinite',
                    },
                    keyframes: {
                        blob: {
                            '0%,100%': { transform: 'translate(0,0) scale(1)' },
                            '33%': { transform: 'translate(30px,-40px) scale(1.1)' },
                            '66%': { transform: 'translate(-20px,20px) scale(0.95)' },
                        },
                        fadeUp: {
                            '0%': { opacity: '0', transform: 'translateY(24px) scale(0.97)' },
                            '100%': { opacity: '1', transform: 'translateY(0) scale(1)' },
                        },
                        shimmer: {
                            '0%': { backgroundPosition: '-200% 0' },
                            '100%': { backgroundPosition: '200% 0' },
                        },
                    },
                }
            }
        }
    </script>
    <style>
        .glass-card {
            background: rgba(255,255,255,0.72);
            backdrop-filter: blur(24px) saturate(180%);
            -webkit-backdrop-filter: blur(24px) saturate(180%);
            border: 1px solid rgba(255,255,255,0.5);
        }
        input:-webkit-autofill,
        input:-webkit-autofill:hover,
        input:-webkit-autofill:focus {
            -webkit-box-shadow: 0 0 0 1000px #fff inset;
            -webkit-text-fill-color: #1e1e1e;
            transition: background-color 5000s ease-in-out 0s;
        }
        @keyframes shake { 0%,100%{transform:translateX(0)} 25%{transform:translateX(-5px)} 75%{transform:translateX(5px)} }
        .shake { animation: shake 0.35s ease-in-out; }
        @keyframes checkPop { 0%{transform:scale(0)} 60%{transform:scale(1.2)} 100%{transform:scale(1)} }
        .check-pop { animation: checkPop 0.4s cubic-bezier(0.16,1,0.3,1); }
        .btn-shimmer {
            background-size: 200% 100%;
            background-image: linear-gradient(90deg, #F85C2C 0%, #FB923C 25%, #F85C2C 50%, #FB923C 75%, #F85C2C 100%);
        }
    </style>
</head>
<body class="min-h-screen flex items-center justify-center p-4 font-sans relative overflow-hidden bg-[#FAFAF9]">

    <!-- Animated gradient background blobs -->
    <div class="fixed inset-0 -z-10">
        <div class="absolute top-[-10%] right-[-5%] w-[500px] h-[500px] bg-brand-200/30 rounded-full blur-3xl animate-blob"></div>
        <div class="absolute bottom-[-15%] left-[-10%] w-[600px] h-[600px] bg-orange-200/25 rounded-full blur-3xl animate-blob" style="animation-delay:-3s"></div>
        <div class="absolute top-[40%] left-[30%] w-[400px] h-[400px] bg-amber-100/30 rounded-full blur-3xl animate-blob" style="animation-delay:-5s"></div>
        <div class="absolute inset-0 bg-gradient-to-br from-white/80 via-transparent to-orange-50/50"></div>
    </div>

    <!-- Main Card -->
    <div class="w-full max-w-[420px] animate-fade-up">
        <div class="glass-card rounded-3xl shadow-[0_8px_40px_rgba(248,92,44,0.10),0_2px_12px_rgba(0,0,0,0.04)] p-8 sm:p-10">

            <!-- Logo & Brand -->
            <div class="text-center mb-6">
                <div class="inline-flex items-center justify-center w-20 h-20 mb-5">
                    <img src="/logo.png" alt="Daraz" class="w-full h-full object-contain">
                </div>
                <h1 class="text-2xl font-bold text-gray-900 tracking-tight">Rider Signup</h1>
                <p class="text-sm text-gray-500 mt-1.5">Create your Barcode Tracker account</p>
            </div>

            <!-- Error Message -->
            <div id="errorMsg" class="hidden mb-5 bg-red-50 border border-red-200 text-red-700 text-sm font-medium px-4 py-3 rounded-xl flex items-start gap-2.5">
                <svg class="w-4 h-4 mt-0.5 shrink-0" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.28 7.22a.75.75 0 00-1.06 1.06L8.94 10l-1.72 1.72a.75.75 0 101.06 1.06L10 11.06l1.72 1.72a.75.75 0 101.06-1.06L11.06 10l1.72-1.72a.75.75 0 00-1.06-1.06L10 8.94 8.28 7.22z" clip-rule="evenodd"/></svg>
                <span id="errorText"></span>
            </div>

            <!-- Signup Form -->
            <form id="signupForm" class="space-y-4">

                <!-- Phone Input -->
                <div class="space-y-1.5">
                    <label for="phone" class="block text-sm font-semibold text-gray-700">Phone Number</label>
                    <div class="relative">
                        <div class="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                            <svg class="w-4.5 h-4.5 text-gray-400" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M2.25 6.75c0 8.284 6.716 15 15 15h2.25a2.25 2.25 0 002.25-2.25v-1.372c0-.516-.351-.966-.852-1.091l-4.423-1.106c-.44-.11-.902.055-1.173.417l-.97 1.293c-.282.376-.769.542-1.21.38a12.035 12.035 0 01-7.143-7.143c-.162-.441.004-.928.38-1.21l1.293-.97c.363-.271.527-.734.417-1.173L6.963 3.102a1.125 1.125 0 00-1.091-.852H4.5A2.25 2.25 0 002.25 4.5v2.25z"/></svg>
                        </div>
                        <input type="tel" id="phone" name="phone" required maxlength="10" placeholder="e.g., 98XXXXXXXX"
                            class="w-full pl-10 pr-4 py-3 bg-white border-2 border-gray-200 rounded-xl text-sm font-medium text-gray-900 placeholder:text-gray-400 focus:outline-none focus:border-brand-500 focus:ring-4 focus:ring-brand-500/10 transition-all duration-200">
                    </div>
                </div>

                <!-- Name Input -->
                <div class="space-y-1.5">
                    <label for="name" class="block text-sm font-semibold text-gray-700">Full Name</label>
                    <div class="relative">
                        <div class="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                            <svg class="w-4.5 h-4.5 text-gray-400" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M15.75 6a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0zM4.501 20.118a7.5 7.5 0 0114.998 0A17.933 17.933 0 0112 21.75c-2.676 0-5.216-.584-7.499-1.632z"/></svg>
                        </div>
                        <input type="text" id="name" name="name" required placeholder="Enter your full name"
                            class="w-full pl-10 pr-4 py-3 bg-white border-2 border-gray-200 rounded-xl text-sm font-medium text-gray-900 placeholder:text-gray-400 focus:outline-none focus:border-brand-500 focus:ring-4 focus:ring-brand-500/10 transition-all duration-200">
                    </div>
                </div>

                <!-- PIN Input -->
                <div class="space-y-1.5">
                    <label for="pin" class="block text-sm font-semibold text-gray-700">PIN</label>
                    <div class="relative">
                        <div class="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                            <svg class="w-4.5 h-4.5 text-gray-400" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M16.5 10.5V6.75a4.5 4.5 0 10-9 0v3.75m-.75 11.25h10.5a2.25 2.25 0 002.25-2.25v-6.75a2.25 2.25 0 00-2.25-2.25H6.75a2.25 2.25 0 00-2.25 2.25v6.75a2.25 2.25 0 002.25 2.25z"/></svg>
                        </div>
                        <input type="password" id="pin" name="pin" required placeholder="Create a PIN (min 4 characters)"
                            class="w-full pl-10 pr-4 py-3 bg-white border-2 border-gray-200 rounded-xl text-sm font-medium text-gray-900 placeholder:text-gray-400 focus:outline-none focus:border-brand-500 focus:ring-4 focus:ring-brand-500/10 transition-all duration-200 tracking-widest">
                    </div>
                </div>

                <!-- Confirm PIN Input -->
                <div class="space-y-1.5">
                    <label for="confirm-pin" class="block text-sm font-semibold text-gray-700">Confirm PIN</label>
                    <div class="relative">
                        <div class="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                            <svg class="w-4.5 h-4.5 text-gray-400" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75m-3-7.036A11.959 11.959 0 013.598 6 11.99 11.99 0 003 9.749c0 5.592 3.824 10.29 9 11.623 5.176-1.332 9-6.03 9-11.622 0-1.31-.21-2.571-.598-3.751h-.152c-3.196 0-6.1-1.248-8.25-3.285z"/></svg>
                        </div>
                        <input type="password" id="confirm-pin" name="confirm-pin" required placeholder="Confirm your PIN"
                            class="w-full pl-10 pr-4 py-3 bg-white border-2 border-gray-200 rounded-xl text-sm font-medium text-gray-900 placeholder:text-gray-400 focus:outline-none focus:border-brand-500 focus:ring-4 focus:ring-brand-500/10 transition-all duration-200 tracking-widest">
                    </div>
                </div>

                <!-- Signup Button -->
                <button type="submit" id="signupBtn"
                    class="w-full py-3.5 px-6 rounded-xl text-white text-sm font-bold tracking-wide uppercase shadow-lg shadow-brand-500/25 hover:shadow-xl hover:shadow-brand-500/30 hover:-translate-y-0.5 active:translate-y-0 transition-all duration-200 btn-shimmer mt-2">
                    Create Account
                </button>
            </form>

            <!-- Login Link -->
            <div class="mt-6 pt-5 border-t border-gray-100/80 text-center">
                <p class="text-sm text-gray-500">Already have an account? <a href="/login" class="font-semibold text-brand-600 hover:text-brand-700 transition-colors">Login here</a></p>
            </div>

            <!-- Footer -->
            <div class="mt-4 text-center">
                <p class="text-xs text-gray-400"><span class="font-semibold text-brand-500">Daraz</span> Nepal &middot; Barcode Operations</p>
            </div>
        </div>
    </div>

    <script>
        function showError(msg) {
            const el = document.getElementById('errorMsg');
            document.getElementById('errorText').textContent = msg;
            el.classList.remove('hidden');
            el.classList.add('shake');
            setTimeout(() => el.classList.remove('shake'), 400);
        }

        function hideError() {
            document.getElementById('errorMsg').classList.add('hidden');
        }

        document.getElementById('signupForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const btn = document.getElementById('signupBtn');
            const phone = document.getElementById('phone').value.trim();
            const name = document.getElementById('name').value.trim();
            const pin = document.getElementById('pin').value;
            const confirmPin = document.getElementById('confirm-pin').value;

            // Validate phone number (exactly 10 digits)
            if (!phone.match(/^\\d{10}$/)) {
                showError('Phone number must be exactly 10 digits');
                return;
            }

            // Validate PIN length
            if (pin.length < 4) {
                showError('PIN must be at least 4 characters');
                return;
            }

            if (pin !== confirmPin) {
                showError('PINs do not match');
                return;
            }

            btn.disabled = true;
            btn.innerHTML = '<span class="inline-flex items-center gap-2"><svg class="animate-spin h-4 w-4" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4" fill="none"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>Creating account...</span>';
            btn.classList.add('opacity-90');
            hideError();

            try {
                const res = await fetch('/api/rider/signup', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ phone: phone, name: name, pin: pin })
                });
                const data = await res.json();
                if (data.ok) {
                    btn.innerHTML = '<span class="inline-flex items-center gap-2 check-pop"><svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5"/></svg>Welcome!</span>';
                    btn.classList.remove('btn-shimmer');
                    btn.style.background = 'linear-gradient(135deg, #22C55E 0%, #16A34A 100%)';
                    btn.style.boxShadow = '0 8px 24px rgba(34,197,94,0.3)';
                    setTimeout(() => { window.location.href = '/'; }, 600);
                } else {
                    showError(data.error || 'Signup failed');
                    btn.disabled = false;
                    btn.innerHTML = 'Create Account';
                    btn.classList.remove('opacity-90');
                }
            } catch (err) {
                showError('Network error');
                btn.disabled = false;
                btn.innerHTML = 'Create Account';
                btn.classList.remove('opacity-90');
            }
        });
    </script>
</body>
</html>
"""

@app.route("/login")
def login_page():
    return redirect(url_for("index"))


@app.route("/signup")
def signup_page():
    return Response(SIGNUP_HTML, content_type="text/html")


@app.route("/api/test", methods=["POST"])
def test_route():
    return jsonify({"ok": True, "message": "Test route works"})


@app.route("/api/rider/login", methods=["POST"])
def rider_login():
    data = request.json
    phone = (data.get("phone") or "").strip()
    pin = (data.get("pin") or "").strip()
    
    if not phone or not pin:
        return jsonify({"ok": False, "error": "Phone and PIN required"}), 400
    
    # Get credentials from sheet
    credentials = get_rider_credentials()
    
    # Check if phone exists and PIN matches
    if phone in credentials and credentials[phone]["password"] == pin:
        session["authenticated"] = True
        session["rider_name"] = credentials[phone]["name"]
        session["phone"] = phone
        return jsonify({"ok": True, "rider_name": credentials[phone]["name"], "phone": phone})
    
    return jsonify({"ok": False, "error": "Invalid phone or PIN"}), 401


@app.route("/api/rider/logout", methods=["POST"])
def rider_logout():
    session.clear()
    return jsonify({"ok": True})


@app.route("/api/rider/session")
def rider_session():
    if session.get("authenticated"):
        return jsonify({"ok": True, "rider_name": session.get("rider_name"), "phone": session.get("phone")})
    return jsonify({"ok": False, "authenticated": False}), 401


@app.route("/api/rider/signup", methods=["POST"])
def rider_signup():
    """Register a new rider with phone, name, and PIN."""
    data = request.json
    phone = (data.get("phone") or "").strip()
    name = (data.get("name") or "").strip()
    pin = (data.get("pin") or "").strip()
    
    if not phone or not name or not pin:
        return jsonify({"ok": False, "error": "Phone, name, and PIN required"}), 400
    
    # Validate phone number (exactly 10 digits for Nepal)
    if not phone.isdigit() or len(phone) != 10:
        return jsonify({"ok": False, "error": "Phone number must be exactly 10 digits"}), 400
    
    # Validate PIN length
    if len(pin) < 4:
        return jsonify({"ok": False, "error": "PIN must be at least 4 characters"}), 400
    
    # Check if phone already exists
    if check_phone_exists(phone):
        return jsonify({"ok": False, "error": "Phone number already registered"}), 409
    
    # Register the rider
    result = register_rider(phone, name, pin)
    if "error" in result:
        return jsonify({"ok": False, "error": result["error"]}), 500
    
    # Auto-login after signup
    session["authenticated"] = True
    session["rider_name"] = name
    session["phone"] = phone
    
    return jsonify({"ok": True, "rider_name": name, "phone": phone, "message": "Registration successful"})


@app.route("/api/admin/records")
def admin_records():
    try:
        records = fetch_all_records()
        return jsonify({"ok": True, "records": records})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/admin/update-status", methods=["POST"])
def admin_update_status():
    data = request.json
    row = data.get("row")
    status = (data.get("status") or "").strip()
    if not row or not status:
        return jsonify({"ok": False, "error": "Missing row or status"}), 400
    try:
        result = update_record_status(row, status)
        return jsonify({"ok": True, "result": result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ─── WhatsApp Routes ──────────────────────────────────────────────────────────

@app.route("/api/admin/whatsapp-status")
def admin_whatsapp_status():
    """Check if WhatsApp Web is connected."""
    if not WHATSAPP_AVAILABLE:
        return jsonify({"ok": False, "error": "WhatsApp module not installed"})
    try:
        logged_in = is_logged_in()
        return jsonify({"ok": True, "logged_in": logged_in})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})


@app.route("/api/admin/send-whatsapp", methods=["POST"])
def admin_send_whatsapp():
    """Send a WhatsApp message to a group."""
    if not WHATSAPP_AVAILABLE:
        return jsonify({"ok": False, "error": "WhatsApp module not installed"}), 500
    data = request.json
    group_name = (data.get("group_name") or "").strip()
    message = (data.get("message") or "").strip()
    if not group_name or not message:
        return jsonify({"ok": False, "error": "Missing group_name or message"}), 400
    try:
        result = send_message_to_group(group_name, message)
        return jsonify(result)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ─── HTML Page (embedded) ────────────────────────────────────────────────────
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Barcode Tracker</title>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
<script>
    tailwind.config = {
        theme: {
            extend: {
                fontFamily: { sans: ['Inter', 'system-ui', 'sans-serif'] },
                colors: {
                    brand: { 50:'#FFF7ED',100:'#FFEDD5',200:'#FDBA74',300:'#FB923C',400:'#F97316',500:'#F85C2C',600:'#EA580C',700:'#C2410C' },
                },
                animation: {
                    'blob': 'blob 8s infinite',
                    'fade-up': 'fadeUp 0.6s cubic-bezier(0.16,1,0.3,1)',
                    'shimmer': 'shimmer 2s linear infinite',
                },
                keyframes: {
                    blob: {
                        '0%,100%': { transform: 'translate(0,0) scale(1)' },
                        '33%': { transform: 'translate(30px,-40px) scale(1.1)' },
                        '66%': { transform: 'translate(-20px,20px) scale(0.95)' },
                    },
                    fadeUp: {
                        '0%': { opacity: '0', transform: 'translateY(24px) scale(0.97)' },
                        '100%': { opacity: '1', transform: 'translateY(0) scale(1)' },
                    },
                    shimmer: {
                        '0%': { backgroundPosition: '-200% 0' },
                        '100%': { backgroundPosition: '200% 0' },
                    },
                },
            }
        }
    }
</script>
<style>
:root {
  --primary: #F85C2C; --primary-hover: #E8531F; --primary-light: rgba(248,92,44,0.12);
  --primary-gradient: linear-gradient(135deg, #F85C2C 0%, #FB923C 100%);
  --success: #22C55E; --success-bg: rgba(34,197,94,0.12);
  --danger: #EF4444; --danger-bg: rgba(239,68,68,0.12);
  --warning: #F59E0B; --info: #3B82F6;
  --radius: 14px; --radius-sm: 10px; --radius-xs: 6px;
  --shadow: 0 2px 8px rgba(0,0,0,0.06); 
  --shadow-md: 0 4px 16px rgba(0,0,0,0.08);
  --shadow-lg: 0 12px 32px rgba(0,0,0,0.12);
  --shadow-glow: 0 0 20px rgba(248,92,44,0.15);
  --transition: 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  --font: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  --mono: 'JetBrains Mono', 'Fira Code', 'Consolas', monospace;
  --sidebar-w: 250px;
  --main-padding: 28px;
  --card-gap: 22px;
  --glass-bg: rgba(255,255,255,0.72);
  --glass-border: rgba(255,255,255,0.5);
  --glass-blur: blur(20px) saturate(180%);
}
[data-theme="dark"] {
  --bg: #0a0a0a; --bg-card: #141414; --bg-elevated: #1c1c1c; --bg-hover: #242424; --bg-input: #161922;
  --border: rgba(255,255,255,0.08); --border-light: rgba(255,255,255,0.12);
  --text: #F5F5F5; --text-secondary: #A3A3A3; --text-muted: #737373;
  --sidebar-bg: rgba(20,20,20,0.95); --sidebar-hover: rgba(255,255,255,0.04);
  --glass-bg: rgba(20,20,20,0.8); --glass-border: rgba(255,255,255,0.1);
}
[data-theme="light"] {
  --bg: #f7f7f7; --bg-card: #FFF; --bg-elevated: #F5F5F5; --bg-hover: #EEE; --bg-input: #F3F4F6;
  --border: rgba(0,0,0,0.06); --border-light: rgba(0,0,0,0.1);
  --text: #171717; --text-secondary: #525252; --text-muted: #737373;
  --sidebar-bg: rgba(255,255,255,0.95); --sidebar-hover: rgba(0,0,0,0.03);
}
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;transition:background var(--transition),color var(--transition);overflow-x:hidden}

/* ── Login Page (Premium Tailwind-enhanced) ─────────────────────────── */
.glass-card{background:rgba(255,255,255,0.72);backdrop-filter:blur(24px) saturate(180%);-webkit-backdrop-filter:blur(24px) saturate(180%);border:1px solid rgba(255,255,255,0.5)}
.gradient-text{background:linear-gradient(135deg,#F85C2C 0%,#FB923C 50%,#F97316 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.btn-shimmer{background-size:200% 100%;background-image:linear-gradient(90deg,#F85C2C 0%,#FB923C 25%,#F85C2C 50%,#FB923C 75%,#F85C2C 100%)}
@keyframes shake{0%,100%{transform:translateX(0)}25%{transform:translateX(-5px)}75%{transform:translateX(5px)}}
.shake{animation:shake 0.35s ease-in-out}
@keyframes checkPop{0%{transform:scale(0)}60%{transform:scale(1.2)}100%{transform:scale(1)}}
.check-pop{animation:checkPop 0.4s cubic-bezier(0.16,1,0.3,1)}
#login-page{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100vh;padding:20px;position:relative;overflow:hidden;font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif}
.login-card{position:relative;z-index:1}

/* ── Custom Dropdown (Premium) ───────────────────────────────────────── */
.custom-dropdown{position:relative;margin-bottom:16px}
.custom-dropdown-label{display:block;font-size:12px;font-weight:600;color:#4B5563;margin-bottom:6px;letter-spacing:0.3px;text-transform:uppercase}
.custom-dropdown-trigger{width:100%;padding:12px 16px;border-radius:12px;border:2px solid #E5E7EB;background:#FAFAFA;color:#1F2937;font-size:14px;font-weight:500;outline:none;cursor:pointer;display:flex;align-items:center;justify-content:space-between;transition:all 0.2s ease;user-select:none}
.custom-dropdown-trigger:hover{border-color:#F85C2C;background:#FFF}
.custom-dropdown-trigger.active{border-color:#F85C2C;box-shadow:0 0 0 4px rgba(248,92,44,0.1);background:#FFF}
.custom-dropdown-trigger .placeholder{color:#9CA3AF;font-weight:400}
.custom-dropdown-trigger .arrow{transition:transform 0.2s ease;color:#9CA3AF}
.custom-dropdown-trigger.active .arrow{transform:rotate(180deg)}
.custom-dropdown-menu{position:absolute;top:calc(100% + 6px);left:0;right:0;background:#FFF;border:1px solid #E5E7EB;border-radius:12px;box-shadow:0 10px 40px rgba(0,0,0,0.12);max-height:220px;overflow-y:auto;z-index:1000;opacity:0;pointer-events:none;transform:translateY(-8px);transition:all 0.2s ease}
.custom-dropdown-menu.show{opacity:1;pointer-events:all;transform:translateY(0)}
.custom-dropdown-item{padding:11px 16px;cursor:pointer;transition:all 0.15s ease;font-size:14px;display:flex;align-items:center;gap:8px;border-bottom:1px solid #F3F4F6}
.custom-dropdown-item:last-child{border-bottom:none}
.custom-dropdown-item:hover{background:#FFF7ED;color:#F85C2C}
.custom-dropdown-item.selected{background:#FFF7ED;color:#F85C2C;font-weight:600}
.custom-dropdown-item .check-icon{opacity:0;margin-left:auto;color:#F85C2C}
.custom-dropdown-item.selected .check-icon{opacity:1}
.custom-dropdown-search{padding:10px 14px;border-bottom:1px solid #F3F4F6;position:sticky;top:0;background:#FFF;z-index:1;border-radius:12px 12px 0 0}
.custom-dropdown-search input{width:100%;padding:8px 12px;border-radius:8px;border:1.5px solid #E5E7EB;background:#FAFAFA;color:#1F2937;font-size:13px;outline:none;transition:all 0.2s ease}
.custom-dropdown-search input:focus{border-color:#F85C2C;box-shadow:0 0 0 3px rgba(248,92,44,0.08)}

/* ── App Layout ─────────────────────────────────────────────────────── */
#app-page{display:none}
.app-layout{display:flex;min-height:100vh}

/* ── Sidebar ───────────────────────────────────────────────────────── */
.sidebar{width:var(--sidebar-w);background:var(--sidebar-bg);backdrop-filter:var(--glass-blur);-webkit-backdrop-filter:var(--glass-blur);border-right:1px solid var(--border);display:flex;flex-direction:column;position:fixed;top:0;left:0;bottom:0;z-index:100;transition:transform .3s cubic-bezier(0.4,0,0.2,1),box-shadow .3s ease;box-shadow:4px 0 24px rgba(0,0,0,0.06)}
.sidebar::before{content:'';position:absolute;top:0;left:0;right:0;width:3px;background:var(--primary-gradient);opacity:0.8}
.sidebar-header{padding:22px 20px;border-bottom:1px solid var(--border);position:relative}
.sidebar-header::after{content:'';position:absolute;bottom:-1px;left:20px;right:20px;height:1px;background:linear-gradient(90deg,transparent,var(--primary-light),transparent)}
.sidebar-header .logo-row{display:flex;align-items:center;gap:12px}
.sidebar-header .logo-icon{width:40px;height:40px;background:var(--primary-gradient);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:18px;box-shadow:0 4px 12px rgba(248,92,44,0.3);transition:transform var(--transition)}
.sidebar-header .logo-icon:hover{transform:scale(1.05) rotate(-2deg)}
.sidebar-header h1{font-size:15px;font-weight:700;background:var(--primary-gradient);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.sidebar-header .rider-badge{font-size:11px;color:var(--text-muted);margin-top:10px;padding:6px 10px;background:var(--bg-elevated);border-radius:var(--radius-xs);border:1px solid var(--border)}
.sidebar-header .rider-badge b{color:var(--primary);font-weight:600}
.sidebar-nav{flex:1;padding:14px 12px;overflow-y:auto}
.sidebar-nav::-webkit-scrollbar{width:4px}
.sidebar-nav::-webkit-scrollbar-track{background:transparent}
.sidebar-nav::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px}
.sidebar-nav::-webkit-scrollbar-thumb:hover{background:var(--text-muted)}
.nav-item{display:flex;align-items:center;gap:12px;padding:11px 14px;border-radius:var(--radius-sm);color:var(--text-secondary);font-size:13px;font-weight:500;cursor:pointer;transition:all var(--transition);margin-bottom:3px;position:relative;overflow:hidden}
.nav-item::before{content:'';position:absolute;inset:0;background:var(--primary-gradient);opacity:0;transition:opacity var(--transition)}
.nav-item:hover{background:var(--sidebar-hover);color:var(--text)}
.nav-item.active{color:var(--primary);font-weight:600}
.nav-item.active::before{opacity:0.08}
.nav-item.active::after{content:'';position:absolute;left:0;top:20%;bottom:20%;width:3px;background:var(--primary-gradient);border-radius:0 3px 3px 0}
.nav-item .icon{font-size:16px;width:20px;text-align:center;position:relative;z-index:1;transition:transform var(--transition)}
.nav-item:hover .icon{transform:scale(1.1)}
.nav-item.active .icon{color:var(--primary)}
.nav-item .badge{margin-left:auto;background:var(--primary-gradient);color:#fff;font-size:10px;padding:3px 8px;border-radius:10px;font-weight:600;position:relative;z-index:1;box-shadow:0 2px 6px rgba(248,92,44,0.3)}
.nav-item span{position:relative;z-index:1}
.sidebar-footer{padding:14px;border-top:1px solid var(--border);position:relative}
.sidebar-footer::before{content:'';position:absolute;top:-1px;left:20px;right:20px;height:1px;background:linear-gradient(90deg,transparent,var(--border),transparent)}
.sidebar-footer .stats-mini{display:flex;gap:8px;margin-bottom:12px}
.stat-mini{flex:1;text-align:center;padding:10px 6px;background:var(--bg-elevated);border-radius:var(--radius-sm);font-size:10px;color:var(--text-muted);border:1px solid var(--border);transition:all var(--transition)}
.stat-mini:hover{border-color:var(--primary);transform:translateY(-1px);box-shadow:var(--shadow)}
.stat-mini b{display:block;font-size:18px;font-weight:700;background:var(--primary-gradient);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.btn-logout{width:100%;padding:10px;background:var(--bg-elevated);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text-secondary);font-size:12px;font-weight:500;cursor:pointer;transition:all var(--transition);display:flex;align-items:center;justify-content:center;gap:6px}
.btn-logout:hover{border-color:var(--danger);color:var(--danger);background:var(--danger-bg);transform:translateY(-1px)}
.theme-toggle{width:100%;padding:10px 12px;background:var(--bg-elevated);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text-secondary);font-size:12px;font-weight:500;cursor:pointer;display:flex;align-items:center;gap:8px;transition:all var(--transition);margin-bottom:8px}
.theme-toggle:hover{border-color:var(--primary);color:var(--primary);background:var(--primary-light);transform:translateY(-1px)}

/* ── Main Content ───────────────────────────────────────────────────── */
.main-content{flex:1;margin-left:var(--sidebar-w);padding:var(--main-padding);max-width:1400px}
.content-grid{display:grid;grid-template-columns:1fr 1fr;gap:var(--card-gap);align-items:start}
.content-grid .col-left{display:flex;flex-direction:column;gap:var(--card-gap)}
.content-grid .col-right{display:flex;flex-direction:column;gap:var(--card-gap)}

/* ── Cards ──────────────────────────────────────────────────────────── */
.card{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:22px;backdrop-filter:var(--glass-blur);-webkit-backdrop-filter:var(--glass-blur);box-shadow:var(--shadow);transition:all var(--transition);position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--primary-gradient);opacity:0;transition:opacity var(--transition)}
.card:hover{box-shadow:var(--shadow-md);transform:translateY(-2px)}
.card:hover::before{opacity:1}
.card-title{font-size:11px;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.8px;margin-bottom:16px;display:flex;align-items:center;gap:8px}
.card-title::before{content:'';width:4px;height:14px;background:var(--primary-gradient);border-radius:2px}

/* ── Scanner ────────────────────────────────────────────────────────── */
#reader{width:100%;border-radius:var(--radius);overflow:hidden;background:#000;box-shadow:var(--shadow-md);border:1px solid var(--border)}
#reader video{border-radius:var(--radius)}
.scan-result{margin-top:14px;padding:14px 18px;border-radius:var(--radius-sm);background:var(--bg-elevated);font-family:var(--mono);font-size:14px;color:var(--text-muted);min-height:48px;display:flex;align-items:center;border:1px solid var(--border);transition:all var(--transition)}
.scan-result.found{border-left:3px solid var(--success);color:var(--text);background:var(--success-bg);border-color:rgba(34,197,94,0.2)}
.scan-confirm-bar{margin-top:14px;padding:18px;border-radius:var(--radius);background:var(--success-bg);border:1px solid rgba(34,197,94,.25);display:none;animation:slideIn .3s cubic-bezier(0.16,1,0.3,1);box-shadow:0 4px 12px rgba(34,197,94,0.1)}
@keyframes slideIn{from{opacity:0;transform:translateY(-10px)}to{opacity:1;transform:translateY(0)}}
.scan-confirm-bar .scan-label{font-size:11px;color:var(--success);font-weight:700;margin-bottom:10px;text-transform:uppercase;letter-spacing:.6px;display:flex;align-items:center;gap:6px}
.scan-confirm-bar .scan-label::before{content:'';width:6px;height:6px;background:var(--success);border-radius:50%;animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}
.scan-confirm-bar .scan-value-display{font-family:var(--mono);font-size:18px;color:var(--text);word-break:break-all;line-height:1.5;padding:6px 0;font-weight:500}
.scan-confirm-bar .scan-edit-input{width:100%;font-family:var(--mono);font-size:16px;color:var(--text);background:var(--bg-card);border:1.5px solid var(--success);border-radius:var(--radius-sm);padding:12px 14px;outline:none;transition:all var(--transition)}
.scan-confirm-bar .scan-edit-input:focus{border-color:var(--success);box-shadow:0 0 0 3px rgba(34,197,94,.15)}
.scan-confirm-actions{display:flex;gap:10px;margin-top:14px}
.scan-confirm-actions .btn{flex:1;padding:12px 16px;font-size:13px;text-align:center}
.scanner-controls{display:flex;gap:10px;margin-top:14px}
.scanner-controls .btn{flex:1}

/* ── Forms ──────────────────────────────────────────────────────────── */
.form-group{margin-bottom:18px}
.form-group label{display:block;font-size:12px;font-weight:600;color:var(--text-secondary);margin-bottom:7px;letter-spacing:0.3px}
.form-group input,.form-group select,.form-group textarea{width:100%;padding:11px 14px;border-radius:var(--radius-sm);border:1.5px solid var(--border);background:var(--bg-elevated);color:var(--text);font-size:14px;outline:none;transition:all var(--transition);font-family:var(--font)}
.form-group input:hover,.form-group select:hover,.form-group textarea:hover{border-color:var(--text-muted)}
.form-group input:focus,.form-group textarea:focus,.form-group select:focus{border-color:var(--primary);box-shadow:0 0 0 3px var(--primary-light);background:var(--bg-card)}
.form-group textarea{resize:vertical;min-height:80px;font-family:var(--font)}
.form-group input::placeholder,.form-group textarea::placeholder{color:var(--text-muted)}

/* ── Buttons ────────────────────────────────────────────────────────── */
.btn{padding:11px 20px;border:none;border-radius:var(--radius-sm);font-size:13px;font-weight:600;cursor:pointer;transition:all var(--transition);display:inline-flex;align-items:center;justify-content:center;gap:7px;position:relative;overflow:hidden}
.btn::before{content:'';position:absolute;inset:0;background:linear-gradient(90deg,transparent,rgba(255,255,255,0.15),transparent);transform:translateX(-100%);transition:transform 0.5s ease}
.btn:hover::before{transform:translateX(100%)}
.btn-primary{background:var(--primary-gradient);color:#fff;box-shadow:0 4px 12px rgba(248,92,44,0.25)}
.btn-primary:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(248,92,44,0.35)}
.btn-primary:active{transform:translateY(0);box-shadow:0 2px 8px rgba(248,92,44,0.25)}
.btn-primary:disabled{opacity:.5;cursor:not-allowed;transform:none;box-shadow:none}
.btn-outline{background:transparent;border:1.5px solid var(--border);color:var(--text-secondary)}
.btn-outline:hover{border-color:var(--primary);color:var(--primary);background:var(--primary-light);transform:translateY(-1px)}
.btn-confirm{background:linear-gradient(135deg,#22C55E 0%,#16A34A 100%);color:#fff;box-shadow:0 4px 12px rgba(34,197,94,0.25)}
.btn-confirm:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(34,197,94,0.35)}
.btn-danger-outline{background:transparent;border:1.5px solid var(--border);color:var(--text-secondary)}
.btn-danger-outline:hover{border-color:var(--danger);color:var(--danger);background:var(--danger-bg);transform:translateY(-1px)}

/* ── Reason Section ─────────────────────────────────────────────────── */
.reason-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px}
.reason-btn{display:flex;flex-direction:column;align-items:center;gap:8px;padding:16px 10px;border-radius:var(--radius);border:1.5px solid var(--border);background:var(--bg-elevated);color:var(--text-secondary);cursor:pointer;transition:all var(--transition);font-size:12px;font-weight:500;text-align:center;position:relative;overflow:hidden}
.reason-btn::before{content:'';position:absolute;inset:0;background:var(--primary-gradient);opacity:0;transition:opacity var(--transition)}
.reason-btn:hover{border-color:var(--primary);color:var(--primary);background:var(--primary-light);transform:translateY(-2px);box-shadow:var(--shadow)}
.reason-btn.selected{border-color:transparent;color:var(--primary);font-weight:600;background:var(--bg-card)}
.reason-btn.selected::before{opacity:0.08}
.reason-btn.selected::after{content:'';position:absolute;inset:0;border-radius:var(--radius);padding:1.5px;background:var(--primary-gradient);-webkit-mask:linear-gradient(#fff 0 0) content-box,linear-gradient(#fff 0 0);-webkit-mask-composite:xor;mask-composite:exclude;pointer-events:none}
.reason-btn .reason-icon{font-size:22px;line-height:1;transition:transform var(--transition)}
.reason-btn:hover .reason-icon{transform:scale(1.15)}
.reason-btn .reason-label{line-height:1.3;position:relative;z-index:1}
.sub-reason-wrap{animation:slideIn .3s cubic-bezier(0.16,1,0.3,1)}
.sub-reason-wrap select{width:100%;padding:11px 14px;border-radius:var(--radius-sm);border:1.5px solid var(--border);background:var(--bg-elevated);color:var(--text);font-size:14px;outline:none;transition:all var(--transition);cursor:pointer;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%23737373' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:36px}
.sub-reason-wrap select:focus{border-color:var(--primary);box-shadow:0 0 0 3px var(--primary-light);background:var(--bg-card)}

/* ── Dashboard Comments ───────────────────────────────────────────── */
.dash-comments-list{max-height:400px;overflow-y:auto}
.dash-comment-item{display:flex;align-items:center;gap:10px;padding:12px 0;border-bottom:1px solid var(--border);cursor:pointer;transition:background var(--transition);flex-wrap:nowrap}
.dash-comment-item:last-child{border-bottom:none}
.dash-comment-item:hover{background:rgba(0,0,0,0.02)}
.dc-tracking-row{font-family:var(--mono);font-size:12px;color:var(--text);word-break:break-all;overflow-wrap:anywhere;flex:1;min-width:0}
.dc-status{font-size:10px;padding:2px 8px;border-radius:10px;font-weight:600;background:var(--primary-light);color:var(--primary)}
.dc-new-dot{width:8px;height:8px;border-radius:50%;background:#22C55E;flex-shrink:0;animation:dcPulse 2s ease-in-out infinite;align-self:center}
.dc-new-badge{font-size:9px;font-weight:700;background:#22C55E;color:#fff;padding:1px 6px;border-radius:8px;flex-shrink:0}
@keyframes dcPulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:0.4;transform:scale(0.8)}}
.dc-comment-entry{display:flex;align-items:flex-start;gap:10px;padding:12px 0;border-bottom:1px solid var(--border)}
.dc-comment-entry:last-child{border-bottom:none}
.dc-comment-ts{color:var(--text-muted);font-size:10px}
.dc-comment-text{color:var(--text-secondary);font-size:12px;line-height:1.5;word-break:break-word}
.dc-empty{color:var(--text-muted);text-align:center;padding:20px;font-size:13px}
.dc-refresh-info{text-align:center;color:var(--text-muted);font-size:10px;margin-top:8px}

/* ── Comment Modal ──────────────────────────────────────────────────── */
.dc-modal-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.6);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);z-index:10000;display:flex;align-items:center;justify-content:center;animation:dcFadeIn 0.25s cubic-bezier(0.16,1,0.3,1);opacity:0;pointer-events:none;transition:opacity 0.25s ease}
.dc-modal-overlay.active{opacity:1;pointer-events:all}
@keyframes dcFadeIn{from{opacity:0}to{opacity:1}}
.dc-modal{background:var(--bg-card);border-radius:var(--radius);width:90%;max-width:500px;max-height:82vh;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,0.2),0 0 0 1px var(--border);border:1px solid var(--border);transform:scale(0.92) translateY(10px);transition:transform 0.3s cubic-bezier(0.16,1,0.3,1);overflow:hidden}
.dc-modal-overlay.active .dc-modal{transform:scale(1) translateY(0)}
.dc-modal-header{display:flex;align-items:center;gap:12px;padding:18px 22px;border-bottom:1px solid var(--border);position:relative}
.dc-modal-header::after{content:'';position:absolute;bottom:-1px;left:22px;right:22px;height:1px;background:linear-gradient(90deg,transparent,var(--primary-light),transparent)}
.dc-modal-tracking{font-family:var(--mono);font-size:13px;font-weight:600;color:var(--text);flex:1;word-break:break-all}
.dc-modal-close{width:30px;height:30px;border-radius:50%;border:none;background:var(--bg-elevated);color:var(--text-secondary);cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all var(--transition)}
.dc-modal-close:hover{background:var(--danger-bg);color:var(--danger);transform:rotate(90deg)}
.dc-modal-body{flex:1;overflow-y:auto;padding:0}
.dc-modal-body::-webkit-scrollbar{width:4px}
.dc-modal-body::-webkit-scrollbar-track{background:transparent}
.dc-modal-body::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px}
.dc-modal-footer{padding:14px 22px;border-top:1px solid var(--border);display:flex;gap:10px;background:var(--bg-elevated)}
.dc-modal-footer input{flex:1;font-size:13px;padding:11px 14px;border-radius:var(--radius-sm);border:1.5px solid var(--border);background:var(--bg-card);color:var(--text);outline:none;font-family:var(--font);transition:all var(--transition)}
.dc-modal-footer input:focus{border-color:var(--primary);box-shadow:0 0 0 3px var(--primary-light)}
.dc-modal-footer button{padding:11px 20px;background:var(--primary-gradient);color:#fff;border:none;border-radius:var(--radius-sm);font-size:13px;font-weight:600;cursor:pointer;transition:all var(--transition);box-shadow:0 2px 8px rgba(248,92,44,0.25)}
.dc-modal-footer button:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(248,92,44,0.35)}
.dc-modal-footer button:disabled{opacity:0.5;cursor:not-allowed;transform:none}
.dc-clickable-tracking{cursor:pointer;transition:all var(--transition);border-radius:4px;padding:2px 4px;margin:-2px -4px}
.dc-clickable-tracking:hover{color:var(--primary);background:var(--primary-light)}

/* Comments Section Header */
.dc-comments-header{display:flex;align-items:center;justify-content:space-between;padding:18px 22px 14px;border-bottom:1px solid var(--border);position:relative}
.dc-comments-header::after{content:'';position:absolute;bottom:-1px;left:22px;right:22px;height:1px;background:linear-gradient(90deg,transparent,var(--primary-light),transparent)}
.dc-comments-header .dc-comments-title{display:flex;align-items:center;gap:8px;font-size:12px;font-weight:700;background:var(--primary-gradient);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;letter-spacing:0.6px;text-transform:uppercase}
.dc-comments-header .dc-comments-title svg{width:16px;height:16px;color:var(--primary)}
.dc-comments-end-btn{display:flex;align-items:center;gap:5px;padding:7px 16px;border-radius:var(--radius-sm);border:none;background:linear-gradient(135deg,#22C55E 0%,#16A34A 100%);color:#fff;font-size:12px;font-weight:600;cursor:pointer;transition:all var(--transition);box-shadow:0 2px 8px rgba(34,197,94,0.25)}
.dc-comments-end-btn:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(34,197,94,0.35)}
.dc-comments-end-btn svg{width:14px;height:14px}
.dc-comments-end-btn:disabled{opacity:0.5;cursor:not-allowed;transform:none;box-shadow:none}

/* Comment Avatar */
.dc-comment-avatar{width:38px;height:38px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:#fff;flex-shrink:0;letter-spacing:0.5px;box-shadow:0 2px 8px rgba(0,0,0,0.15)}
.dc-comment-avatar.staff{background:linear-gradient(135deg,#F85C2C 0%,#FB923C 100%)}
.dc-comment-avatar.rider{background:linear-gradient(135deg,#22C55E 0%,#16A34A 100%)}
.dc-comment-avatar.seller{background:linear-gradient(135deg,#3B82F6 0%,#2563EB 100%)}
.dc-comment-content{flex:1;min-width:0}
.dc-comment-author-name{font-size:13px;font-weight:700;color:var(--text);margin-bottom:3px}
.dc-comment-text{font-size:13px;color:var(--text);line-height:1.5}
.dc-comment-timestamp{display:flex;align-items:center;gap:4px;margin-top:6px;font-size:11px;color:var(--text-muted)}
.dc-comment-timestamp svg{width:12px;height:12px;flex-shrink:0}
.dc-comment-entry{display:flex;align-items:flex-start;gap:12px;padding:16px 22px;border-bottom:1px solid var(--border);transition:all var(--transition)}
.dc-comment-entry:hover{background:var(--bg-elevated)}
.dc-comment-entry:last-child{border-bottom:none}

/* ── History ────────────────────────────────────────────────────────── */
.history-list{max-height:520px;overflow-y:auto;padding-right:4px}
.history-list::-webkit-scrollbar{width:4px}
.history-list::-webkit-scrollbar-track{background:transparent}
.history-list::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px}
.history-list::-webkit-scrollbar-thumb:hover{background:var(--text-muted)}
.history-item{display:flex;flex-direction:column;gap:8px;padding:14px 16px;border-bottom:1px solid var(--border);transition:all var(--transition);border-radius:var(--radius-sm);margin-bottom:4px}
.history-item:hover{background:var(--bg-elevated);transform:translateX(2px)}
.history-item:last-child{border-bottom:none}
.h-head{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.h-head .h-time{margin-left:auto}
.h-tracking{font-family:var(--mono);font-size:13px;line-height:1.5;color:var(--text);word-break:normal;overflow-wrap:anywhere;letter-spacing:.3px;font-weight:500}
.h-sub{font-size:11px;color:var(--text-muted);line-height:1.4;display:flex;align-items:center;gap:6px}
.h-time{font-size:11px;color:var(--text-muted);white-space:nowrap;padding:3px 8px;background:var(--bg-elevated);border-radius:var(--radius-xs);border:1px solid var(--border)}
.h-status{font-size:11px;padding:4px 12px;border-radius:12px;font-weight:600;background:var(--bg-elevated);color:var(--text-secondary);white-space:nowrap;border:1px solid var(--border);transition:all var(--transition)}
.h-reason{background:var(--primary-light);color:var(--primary);border-color:rgba(248,92,44,0.2)}
.h-done{background:var(--success-bg);color:var(--success);border-color:rgba(34,197,94,0.2)}

/* ── Toast ──────────────────────────────────────────────────────────── */
.toast{position:fixed;bottom:28px;right:28px;padding:16px 24px;border-radius:var(--radius);font-size:13px;font-weight:600;z-index:10001;opacity:0;transform:translateY(20px) scale(0.95);transition:all .35s cubic-bezier(0.16,1,0.3,1);pointer-events:none;display:flex;align-items:center;gap:12px;backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);border:1px solid rgba(255,255,255,0.1)}
.toast.show{opacity:1;transform:translateY(0) scale(1)}
.toast.success{background:linear-gradient(135deg,#22C55E 0%,#16A34A 100%);color:#fff;box-shadow:0 8px 24px rgba(34,197,94,0.3)}
.toast.error{background:linear-gradient(135deg,#EF4444 0%,#DC2626 100%);color:#fff;box-shadow:0 8px 24px rgba(239,68,68,0.3)}
.toast.info{background:var(--primary-gradient);color:#fff;box-shadow:0 8px 24px rgba(248,92,44,0.3)}

/* ── Mobile Header ──────────────────────────────────────────────────── */
.mobile-header{display:none;position:fixed;top:0;left:0;right:0;height:58px;background:var(--sidebar-bg);backdrop-filter:var(--glass-blur);-webkit-backdrop-filter:var(--glass-blur);border-bottom:1px solid var(--border);padding:0 16px;align-items:center;gap:12px;z-index:99;box-shadow:0 2px 12px rgba(0,0,0,0.06)}
.mobile-header::after{content:'';position:absolute;bottom:-1px;left:0;right:0;height:2px;background:var(--primary-gradient);opacity:0.6}
.hamburger{width:38px;height:38px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:5px;cursor:pointer;border-radius:var(--radius-sm);transition:all var(--transition)}
.hamburger:hover{background:var(--bg-hover);transform:scale(1.05)}
.hamburger:active{transform:scale(0.95)}
.hamburger span{width:18px;height:2px;background:var(--text);border-radius:2px;transition:all var(--transition)}
.mobile-header h1{font-size:15px;font-weight:700;background:var(--primary-gradient);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.sidebar-overlay{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.6);backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);z-index:98;animation:fadeIn 0.2s ease}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}

/* ── Mobile Bottom Navigation ───────────────────────────────────────── */
.mobile-bottom-nav{display:none;position:fixed;bottom:0;left:0;right:0;height:64px;background:var(--sidebar-bg);backdrop-filter:var(--glass-blur);-webkit-backdrop-filter:var(--glass-blur);border-top:1px solid var(--border);z-index:99;box-shadow:0 -2px 12px rgba(0,0,0,0.06)}
.mobile-bottom-nav::before{content:'';position:absolute;top:-1px;left:0;right:0;height:2px;background:var(--primary-gradient);opacity:0.6}
.bottom-nav-items{display:flex;height:100%;align-items:center;justify-content:space-around;padding:0 8px}
.bottom-nav-item{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;flex:1;max-width:80px;height:100%;cursor:pointer;transition:all var(--transition);position:relative;color:var(--text-muted)}
.bottom-nav-item:hover{color:var(--text-secondary)}
.bottom-nav-item.active{color:var(--primary)}
.bottom-nav-item.active::before{content:'';position:absolute;top:0;left:50%;transform:translateX(-50%);width:40px;height:3px;background:var(--primary-gradient);border-radius:0 0 3px 3px}
.bottom-nav-icon{font-size:20px;line-height:1;transition:transform var(--transition)}
.bottom-nav-item.active .bottom-nav-icon{transform:scale(1.1)}
.bottom-nav-label{font-size:10px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%}
.bottom-nav-badge{position:absolute;top:6px;right:calc(50% - 18px);min-width:16px;height:16px;padding:0 4px;background:linear-gradient(135deg,#EF4444,#DC2626);color:#fff;font-size:9px;font-weight:700;border-radius:8px;display:flex;align-items:center;justify-content:center;box-shadow:0 2px 4px rgba(239,68,68,0.3)}

/* ── Responsive: Laptop (>1024px) ───────────────────────────────────── */
@media(min-width:1025px){
  .content-grid{grid-template-columns:1fr 380px}
}

/* ─ Responsive: Tablet (768–1024px) ────────────────────────────────── */
@media(max-width:1024px){
  :root{--sidebar-w:220px;--main-padding:20px;--card-gap:16px}
  .content-grid{grid-template-columns:1fr}
  .reason-grid{grid-template-columns:repeat(3,1fr)}
}

/* ── Responsive: Mobile (<768px) ────────────────────────────────────── */
@media(max-width:768px){
  :root{--sidebar-w:260px;--main-padding:16px;--card-gap:14px}
  .mobile-header{display:flex}
  .mobile-bottom-nav{display:block}
  .sidebar{transform:translateX(-100%);top:0}
  .sidebar.open{transform:translateX(0)}
  .sidebar-overlay.show{display:block}
  .main-content{margin-left:0;padding:68px 14px 78px}
  .content-grid{grid-template-columns:1fr}
  .card{padding:16px;overflow:visible}
  .reason-grid{grid-template-columns:repeat(2,1fr);gap:8px}
  .reason-btn{padding:12px 6px;font-size:11px}
  .reason-btn .reason-icon{font-size:18px}
  .history-list{max-height:350px}
  .scan-confirm-actions{flex-wrap:wrap}
  .scan-confirm-actions .btn{min-width:calc(50% - 4px)}
  .toast{left:16px;right:16px;bottom:78px}
  .custom-dropdown-menu{max-height:200px;position:absolute;z-index:9999}
  .custom-dropdown{position:relative;z-index:100}
}

/* ── Responsive: Small Phone (<400px) ───────────────────────────────── */
@media(max-width:400px){
  :root{--main-padding:12px;--card-gap:12px}
  .card{padding:14px}
  .reason-grid{grid-template-columns:repeat(2,1fr);gap:6px}
  .reason-btn{padding:10px 4px;font-size:10px;gap:4px}
  .reason-btn .reason-icon{font-size:16px}
  .reason-btn .reason-label{font-size:10px}
  .form-group input,.form-group select,.form-group textarea{padding:9px 12px;font-size:13px}
  .btn{padding:9px 14px;font-size:12px}
  .sidebar-header{padding:16px}
  .sidebar-nav{padding:8px}
  .nav-item{padding:9px 12px;font-size:12px}
  .login-card{padding:28px 20px}
}
</style>
</head>
<body>
<div id="login-page">
  <!-- Animated gradient background blobs -->
  <div class="fixed inset-0 -z-10">
    <div class="absolute top-[-10%] right-[-5%] w-[500px] h-[500px] bg-brand-200/30 rounded-full blur-3xl animate-blob"></div>
    <div class="absolute bottom-[-15%] left-[-10%] w-[600px] h-[600px] bg-orange-200/25 rounded-full blur-3xl animate-blob" style="animation-delay:-3s"></div>
    <div class="absolute top-[40%] left-[30%] w-[400px] h-[400px] bg-amber-100/30 rounded-full blur-3xl animate-blob" style="animation-delay:-5s"></div>
    <div class="absolute inset-0 bg-gradient-to-br from-white/80 via-transparent to-orange-50/50"></div>
  </div>

  <!-- Main Card -->
  <div class="w-full max-w-[420px] animate-fade-up">
    <div class="glass-card rounded-3xl shadow-[0_8px_40px_rgba(248,92,44,0.10),0_2px_12px_rgba(0,0,0,0.04)] p-8 sm:p-10 login-card">

      <!-- Logo & Brand -->
      <div class="text-center mb-6">
        <div class="inline-flex items-center justify-center w-20 h-20 mb-5">
          <img src="/logo.png" alt="Daraz" class="w-full h-full object-contain">
        </div>
        <h1 class="text-2xl font-bold text-gray-900 tracking-tight">Barcode Tracker</h1>
        <p class="text-sm text-gray-500 mt-1.5">Select your hub and rider to begin</p>
      </div>

      <!-- Status Badge -->
      <div class="flex items-center justify-center gap-2 bg-emerald-50 border border-emerald-200/60 rounded-full py-2 px-4 mb-6 w-fit mx-auto">
        <span class="relative flex h-2 w-2">
          <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
          <span class="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
        </span>
        <span class="text-xs font-semibold text-emerald-700">Ready to Scan</span>
      </div>

      <!-- Error Message -->
      <div id="login-error" class="hidden mb-4 bg-red-50 border border-red-200 text-red-700 text-sm font-medium px-4 py-3 rounded-xl flex items-start gap-2.5">
        <svg class="w-4 h-4 mt-0.5 shrink-0" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.28 7.22a.75.75 0 00-1.06 1.06L8.94 10l-1.72 1.72a.75.75 0 101.06 1.06L10 11.06l1.72 1.72a.75.75 0 101.06-1.06L11.06 10l1.72-1.72a.75.75 0 00-1.06-1.06L10 8.94 8.28 7.22z" clip-rule="evenodd"/></svg>
        <span id="login-error-text"></span>
      </div>

      <!-- Login Form -->
      <div id="login-form" style="display:none" class="space-y-1">
        <div class="custom-dropdown">
          <span class="custom-dropdown-label">Hub</span>
          <div class="custom-dropdown-trigger" id="hub-trigger" onclick="toggleDropdown('hub')">
            <span class="placeholder" id="hub-placeholder">-- Select Hub --</span>
            <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:16px;height:16px"><polyline points="6 9 12 15 18 9"/></svg>
          </div>
          <div class="custom-dropdown-menu" id="hub-menu">
            <div class="custom-dropdown-search"><input type="text" placeholder="Search hubs..." oninput="filterDropdown('hub', this.value)" onclick="event.stopPropagation()"></div>
            <div id="hub-options"></div>
          </div>
          <input type="hidden" id="login-hub" value="">
        </div>

        <div class="custom-dropdown">
          <span class="custom-dropdown-label">Rider Name</span>
          <div class="custom-dropdown-trigger" id="rider-trigger" onclick="toggleDropdown('rider')">
            <span class="placeholder" id="rider-placeholder">-- Select Rider --</span>
            <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:16px;height:16px"><polyline points="6 9 12 15 18 9"/></svg>
          </div>
          <div class="custom-dropdown-menu" id="rider-menu">
            <div class="custom-dropdown-search"><input type="text" placeholder="Search riders..." oninput="filterDropdown('rider', this.value)" onclick="event.stopPropagation()"></div>
            <div id="rider-options"></div>
          </div>
          <input type="hidden" id="login-rider" value="">
        </div>

        <button class="w-full py-3.5 px-6 rounded-xl text-white text-sm font-bold tracking-wide uppercase shadow-lg shadow-brand-500/25 hover:shadow-xl hover:shadow-brand-500/30 hover:-translate-y-0.5 active:translate-y-0 transition-all duration-200 btn-shimmer mt-3 border-0 cursor-pointer" onclick="doLogin()">
          Start Scanning
        </button>
      </div>

      <!-- Loading State -->
      <div id="login-loading" class="text-center py-6" style="display:none">
        <div class="inline-flex items-center gap-3 text-gray-500 text-sm font-medium">
          <svg class="animate-spin h-5 w-5 text-brand-500" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4" fill="none"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>
          Loading riders...
        </div>
      </div>

      <!-- Load Error State -->
      <div id="login-load-error" class="text-center text-red-600 text-sm font-medium py-4" style="display:none"></div>
      <button id="login-retry" style="display:none" class="w-full py-3 px-6 rounded-xl border-2 border-gray-200 bg-transparent text-gray-600 text-sm font-semibold cursor-pointer hover:border-brand-500 hover:text-brand-600 transition-all duration-200 mt-2" onclick="loadRiders()">Retry</button>

      <!-- Footer -->
      <div class="mt-6 pt-5 border-t border-gray-100/80 text-center">
        <p class="text-xs text-gray-400"><span class="font-semibold text-brand-500">Daraz</span> Nepal &middot; Barcode Operations</p>
      </div>
    </div>
  </div>
</div>
<div id="app-page">
<div class="mobile-header"><div class="hamburger" onclick="toggleSidebar()"><span></span><span></span><span></span></div><h1>Barcode Tracker</h1></div>
<div class="sidebar-overlay" id="sidebar-overlay" onclick="toggleSidebar()"></div>
<div class="app-layout">
  <aside class="sidebar" id="sidebar">
    <div class="sidebar-header">
      <div class="logo-row"><div class="logo-icon">&#128230;</div><h1>Barcode Tracker</h1></div>
      <div class="rider-badge"><b id="hdr-rider">-</b> &middot; <span id="hdr-hub">-</span></div>
    </div>
    <nav class="sidebar-nav">
      <div class="nav-item active" onclick="scrollToSection('scanner-section',this)"><span class="icon">&#128247;</span> Scanner</div>
      <div class="nav-item" onclick="scrollToSection('reason-section',this)"><span class="icon">&#9888;</span> Reason &amp; Submit</div>
      <div class="nav-item" onclick="scrollToSection('history-section',this)"><span class="icon">&#128337;</span> Recent Scans <span class="badge" id="nav-history-count">0</span></div>
      <div class="nav-item" onclick="scrollToSection('comments-section',this)"><span class="icon">&#128172;</span> Comments <span class="badge" id="nav-comments-count">0</span></div>
    </nav>
    <div class="sidebar-footer">
      <div class="stats-mini">
        <div class="stat-mini"><b id="s-total">0</b>Total</div>
        <div class="stat-mini"><b id="s-synced">0</b>Synced</div>
        <div class="stat-mini"><b id="s-pending">0</b>Pending</div>
      </div>
      <button class="theme-toggle" onclick="toggleTheme()"><span id="theme-icon">&#9728;</span> <span id="theme-label">Dark Mode</span></button>
      <button class="btn-logout" style="margin-top:8px" onclick="doLogout()">Switch Rider</button>
    </div>
  </aside>
  <main class="main-content">
    <div class="content-grid">
      <div class="col-left">
        <div class="card" id="scanner-section">
          <div class="card-title">Scanner</div>
          <div id="reader"></div>
          <div id="scan-result" class="scan-result">Waiting for barcode...</div>
          <div id="scan-confirm" class="scan-confirm-bar">
            <div class="scan-label">Scanned Barcode</div>
            <div class="scan-value-display" id="scan-display"></div>
            <input class="scan-edit-input" id="scan-edit" placeholder="Edit tracking number...">
            <div class="scan-confirm-actions">
              <button class="btn btn-confirm" onclick="confirmScan()">Confirm</button>
              <button class="btn btn-outline" id="btn-edit-scan" onclick="toggleEditScan()">Edit</button>
              <button class="btn btn-danger-outline" onclick="rescanBarcode()">Rescan</button>
            </div>
          </div>
          <div class="scanner-controls">
            <button class="btn btn-outline" id="btn-start" onclick="startScanner()">Start Camera</button>
            <button class="btn btn-outline" id="btn-stop" onclick="stopScanner()" style="display:none">Stop</button>
          </div>
        </div>
        <div class="card" id="reason-section">
          <div class="card-title">Reason</div>
          <div class="form-group"><label>Tracking Number</label><input type="text" id="f-tracking" placeholder="Scanned or type manually" style="font-family:var(--mono)"></div>
          <div class="reason-grid" id="reason-grid">
            <button class="reason-btn" data-category="Call Issue" onclick="selectReason(this)">
              <span class="reason-icon">&#128222;</span>
              <span class="reason-label">Call Issue</span>
            </button>
            <button class="reason-btn" data-category="Customer Issue" onclick="selectReason(this)">
              <span class="reason-icon">&#128100;</span>
              <span class="reason-label">Customer Issue</span>
            </button>
            <button class="reason-btn" data-category="Product / Address Issues" onclick="selectReason(this)">
              <span class="reason-icon">&#128230;</span>
              <span class="reason-label">Product / Address Issues</span>
            </button>
            <button class="reason-btn" data-category="Operational Issue" onclick="selectReason(this)">
              <span class="reason-icon">&#128736;</span>
              <span class="reason-label">Operational Issue</span>
            </button>
            <button class="reason-btn" data-category="Other" onclick="selectReason(this)">
              <span class="reason-icon">&#8943;</span>
              <span class="reason-label">Other</span>
            </button>
          </div>
          <div class="sub-reason-wrap" id="sub-reason-wrap" style="display:none">
            <div class="custom-dropdown" style="margin-bottom:0">
              <span class="custom-dropdown-label" id="sub-reason-label">Sub-Reason</span>
              <div class="custom-dropdown-trigger" id="sub-reason-trigger" onclick="toggleDropdown('sub-reason')">
                <span class="placeholder" id="sub-reason-placeholder">-- Select reason --</span>
                <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:16px;height:16px"><polyline points="6 9 12 15 18 9"/></svg>
              </div>
              <div class="custom-dropdown-menu" id="sub-reason-menu">
                <div class="custom-dropdown-search"><input type="text" placeholder="Search reasons..." oninput="filterDropdown('sub-reason', this.value)" onclick="event.stopPropagation()"></div>
                <div id="sub-reason-options"></div>
              </div>
              <input type="hidden" id="f-sub-reason" value="">
            </div>
            <div class="form-group" id="custom-reason-wrap" style="display:none;margin-top:12px;margin-bottom:0">
              <label>Custom Reason</label>
              <input type="text" id="f-custom-reason" placeholder="Type your reason...">
            </div>
          </div>
          <div style="margin-top:16px">
            <div class="form-group"><label>Remarks (optional)</label><textarea id="f-notes" placeholder="Any additional details..." rows="2"></textarea></div>
            <button class="btn btn-primary" style="width:100%;padding:12px" id="btn-submit" onclick="submitRecord()">Submit</button>
          </div>
        </div>
      </div>
      <div class="col-right">
        <div class="card" id="history-section">
          <div class="card-title">Recent Scans</div>
          <div class="history-list" id="history"></div>
        </div>
        <div class="card" id="comments-section">
          <div class="card-title">Dashboard Comments</div>
          <div class="dash-comments-list" id="dash-comments">
            <div class="dc-empty">Loading comments...</div>
          </div>
          <div class="dc-refresh-info">Click a tracking number to view/add comments. Auto-refreshes every 10s</div>
        </div>
      </div>
    </div>
  </main>
</div>
</div>

<!-- Mobile Bottom Navigation -->
<nav class="mobile-bottom-nav" id="mobile-bottom-nav">
  <div class="bottom-nav-items">
    <div class="bottom-nav-item active" data-section="scanner-section" onclick="switchBottomTab(this)">
      <span class="bottom-nav-icon">&#128247;</span>
      <span class="bottom-nav-label">Scanner</span>
    </div>
    <div class="bottom-nav-item" data-section="reason-section" onclick="switchBottomTab(this)">
      <span class="bottom-nav-icon">&#9888;</span>
      <span class="bottom-nav-label">Reason</span>
    </div>
    <div class="bottom-nav-item" data-section="history-section" onclick="switchBottomTab(this)">
      <span class="bottom-nav-icon">&#128337;</span>
      <span class="bottom-nav-label">Recent</span>
      <span class="bottom-nav-badge" id="bottom-history-count" style="display:none">0</span>
    </div>
    <div class="bottom-nav-item" data-section="comments-section" onclick="switchBottomTab(this)">
      <span class="bottom-nav-icon">&#128172;</span>
      <span class="bottom-nav-label">Comments</span>
      <span class="bottom-nav-badge" id="bottom-comments-count" style="display:none">0</span>
    </div>
  </div>
</nav>

<div class="toast" id="toast"></div>

<!-- Comment Modal -->
<div class="dc-modal-overlay" id="dc-modal-overlay" style="display:none">
  <div class="dc-modal">
    <div class="dc-modal-header">
      <span class="dc-modal-tracking" id="dc-modal-tracking">Tracking Number</span>
      <button class="dc-modal-close" onclick="closeRiderCommentModal()">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </button>
    </div>
    <div class="dc-comments-header">
      <div class="dc-comments-title">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/></svg>
        COMMENTS
      </div>
      <button class="dc-comments-end-btn" id="dc-modal-end-btn" onclick="endPackageJourney()">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>
        End
      </button>
    </div>
    <div class="dc-modal-body" id="dc-modal-body">
      <div class="dc-empty">Loading comments...</div>
    </div>
    <div class="dc-modal-footer">
      <input type="text" id="dc-modal-input" placeholder="Type your reply..." onkeydown="if(event.key==='Enter')submitRiderCommentFromModal()">
      <button id="dc-modal-submit" onclick="submitRiderCommentFromModal()">Send</button>
    </div>
  </div>
</div>

<script>
let html5QrCode=null,lastScan=null,history=JSON.parse(localStorage.getItem('bt_history')||'[]');
let riderHubs={},currentRider=null,currentHub=null;
let selectedReasonCategory='',selectedSubReason='';
const REASON_SUBS={
  'Call Issue':['Phone switch-off','Phone not connected','Phone number wrong','Phone cut by customer','Phone not answered by customer'],
  'Customer Issue':['Customer phone answered by family/others','Customer not at location due to office hours ends','Customer cancelled order before delivery','Customer did not placed an order','Customer only said cancelled the package','Customer has no cash available','Customer found same product cheaper elsewhere','Customer got confused about order','Customer placed the same order on other platforms','Customer want to open package','Customer want package on urgent basis','Customer claims delay delivery','Customer not at home'],
  'Product / Address Issues':['Duplicated order','COD amount wrong','Invalid address','Non serviceable area','Item not as Photo','Address changed to another location','Package damaged','Size/color issue'],
  'Operational Issue':['Driver overloaded','Vehicle breakdown / battery swipe issue','Rider mobile issue','Rider health issue','Urgent leave required to rider'],
  'Other':[]
};
function toggleTheme(){const h=document.documentElement,c=h.getAttribute('data-theme'),n=c==='dark'?'light':'dark';h.setAttribute('data-theme',n);document.getElementById('theme-icon').innerHTML=n==='dark'?'&#9790;':'&#9728;';document.getElementById('theme-label').textContent=n==='dark'?'Light Mode':'Dark Mode'}
function toggleSidebar(){document.getElementById('sidebar').classList.toggle('open');document.getElementById('sidebar-overlay').classList.toggle('show')}
function scrollToSection(id,el){const s=document.getElementById(id);if(s)s.scrollIntoView({behavior:'smooth',block:'start'});document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));if(el)el.classList.add('active');if(window.innerWidth<=768)toggleSidebar();if(id==='history-section'){markAllScansSeen();document.getElementById('nav-history-count').textContent='0';updateBottomBadges()}else if(id==='comments-section'){if(_lastCommentRecords.length>0){_lastCommentRecords.forEach(r=>{markSeen(r.tracking,r.row,dcHash(r.comments||''))});document.getElementById('nav-comments-count').textContent='0';updateBottomBadges()}}}
function switchBottomTab(el){const section=el.getAttribute('data-section');document.querySelectorAll('.bottom-nav-item').forEach(n=>n.classList.remove('active'));el.classList.add('active');const s=document.getElementById(section);if(s)s.scrollIntoView({behavior:'smooth',block:'start'});document.querySelectorAll('.nav-item').forEach(n=>{n.classList.remove('active');if(n.getAttribute('onclick')&&n.getAttribute('onclick').includes(section))n.classList.add('active')});if(section==='history-section'){markAllScansSeen();document.getElementById('nav-history-count').textContent='0';updateBottomBadges()}else if(section==='comments-section'){if(_lastCommentRecords.length>0){_lastCommentRecords.forEach(r=>{markSeen(r.tracking,r.row,dcHash(r.comments||''))});document.getElementById('nav-comments-count').textContent='0';updateBottomBadges()}}}
function updateBottomBadges(){const hc=document.getElementById('nav-history-count');const cc=document.getElementById('nav-comments-count');const bh=document.getElementById('bottom-history-count');const bc=document.getElementById('bottom-comments-count');if(hc&&bh){const v=parseInt(hc.textContent)||0;bh.textContent=v;bh.style.display=v>0?'flex':'none'}if(cc&&bc){const v=parseInt(cc.textContent)||0;bc.textContent=v;bc.style.display=v>0?'flex':'none'}}
async function loadRiders(){document.getElementById('login-loading').style.display='';document.getElementById('login-form').style.display='none';document.getElementById('login-load-error').style.display='none';document.getElementById('login-retry').style.display='none';try{const r=await fetch('/api/riders'),d=await r.json();if(!d.ok)throw new Error(d.error||'Failed to load riders');riderHubs=d.hubs||{};const h=Object.keys(riderHubs).sort();if(h.length===0)throw new Error('No active riders found.');renderDropdownOptions('hub',h);document.getElementById('login-loading').style.display='none';document.getElementById('login-form').style.display='';document.getElementById('login-load-error').style.display='none'}catch(e){document.getElementById('login-loading').style.display='none';document.getElementById('login-form').style.display='';document.getElementById('login-load-error').innerHTML='<span class="inline-flex items-center gap-1.5 text-amber-700"><svg class="w-4 h-4" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M8.485 2.495c.673-1.167 2.357-1.167 3.03 0l6.28 10.875c.673 1.167-.17 2.625-1.516 2.625H3.72c-1.347 0-2.189-1.458-1.515-2.625L8.485 2.495zM10 6a.75.75 0 01.75.75v3.5a.75.75 0 01-1.5 0v-3.5A.75.75 0 0110 6zm0 9a1 1 0 100-2 1 1 0 000 2z" clip-rule="evenodd"/></svg>Could not connect to data server. Dropdowns may be empty.</span>';document.getElementById('login-load-error').style.display='';document.getElementById('login-retry').style.display=''}}

// Custom Dropdown Functions
function renderDropdownOptions(type, items){
  const container = document.getElementById(type + '-options');
  container.innerHTML = items.map(item => 
    '<div class="custom-dropdown-item" data-value="' + item + '" onclick="selectDropdownItem(\'' + type + '\', \'' + item.replace(/'/g, "\\'") + '\')">' +
    '<span>' + item + '</span>' +
    '<svg class="check-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:16px;height:16px"><polyline points="20 6 9 17 4 12"/></svg>' +
    '</div>'
  ).join('');
}

function toggleDropdown(type){
  const trigger = document.getElementById(type + '-trigger');
  const menu = document.getElementById(type + '-menu');
  const isActive = trigger.classList.contains('active');
  
  // Close all dropdowns first
  document.querySelectorAll('.custom-dropdown-trigger').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.custom-dropdown-menu').forEach(m => m.classList.remove('show'));
  
  if(!isActive){
    trigger.classList.add('active');
    menu.classList.add('show');
    const searchInput = menu.querySelector('input');
    if(searchInput) setTimeout(() => searchInput.focus(), 100);
  }
}

function selectDropdownItem(type, value){
  // Handle different hidden input IDs
  const hiddenId = type === 'sub-reason' ? 'f-sub-reason' : 'login-' + type;
  document.getElementById(hiddenId).value = value;
  document.getElementById(type + '-placeholder').textContent = value;
  document.getElementById(type + '-placeholder').classList.remove('placeholder');
  
  // Update selected state
  const container = document.getElementById(type + '-options');
  container.querySelectorAll('.custom-dropdown-item').forEach(item => {
    item.classList.toggle('selected', item.getAttribute('data-value') === value);
  });
  
  // Close dropdown
  document.getElementById(type + '-trigger').classList.remove('active');
  document.getElementById(type + '-menu').classList.remove('show');
  
  // Update global variable for sub-reason
  if(type === 'sub-reason'){
    selectedSubReason = value;
    onSubReasonChange();
  }
  
  // If hub selected, update rider options
  if(type === 'hub'){
    const riders = riderHubs[value] || [];
    renderDropdownOptions('rider', riders);
    // Reset rider selection
    document.getElementById('login-rider').value = '';
    document.getElementById('rider-placeholder').textContent = '-- Select Rider --';
    document.getElementById('rider-placeholder').classList.add('placeholder');
    document.getElementById('rider-options').querySelectorAll('.custom-dropdown-item').forEach(item => item.classList.remove('selected'));
  }
}

function filterDropdown(type, query){
  const container = document.getElementById(type + '-options');
  const items = container.querySelectorAll('.custom-dropdown-item');
  const lowerQuery = query.toLowerCase();
  items.forEach(item => {
    const value = item.getAttribute('data-value').toLowerCase();
    item.style.display = value.includes(lowerQuery) ? '' : 'none';
  });
}

// Close dropdowns when clicking outside
document.addEventListener('click', function(e){
  if(!e.target.closest('.custom-dropdown')){
    document.querySelectorAll('.custom-dropdown-trigger').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.custom-dropdown-menu').forEach(m => m.classList.remove('show'));
  }
});
function doLogin(){const h=document.getElementById('login-hub').value,r=document.getElementById('login-rider').value;if(!h||!r){document.getElementById('login-error-text').textContent='Please select both hub and rider.';document.getElementById('login-error').classList.remove('hidden');document.getElementById('login-error').classList.add('shake');setTimeout(()=>document.getElementById('login-error').classList.remove('shake'),400);return}currentRider=r;currentHub=h;loadSeenState(r);_lastCommentsHash='';_lastCommentRecords=[];document.getElementById('login-page').style.display='none';document.getElementById('app-page').style.display='block';document.getElementById('hdr-rider').textContent=r;document.getElementById('hdr-hub').textContent=h;renderHistory();startCommentRefresh()}
function doLogout(){stopScanner();stopCommentRefresh();currentRider=null;currentHub=null;_lastCommentsHash='';_lastCommentRecords=[];document.getElementById('app-page').style.display='none';document.getElementById('login-page').style.display='';loadRiders()}
/* The QR scanner library is fetched from a CDN only when the rider actually
   starts the camera. Keeping it out of <head> means a slow or unreachable
   CDN can never leave the rider staring at a blank screen. */
let _qrLibPromise=null;
function loadQrLibrary(){
  if(window.Html5Qrcode)return Promise.resolve(true);
  if(!_qrLibPromise){
    _qrLibPromise=new Promise(function(resolve){
      var s=document.createElement('script');
      s.src='https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js';
      s.onload=function(){resolve(true)};
      s.onerror=function(){resolve(false)};
      document.head.appendChild(s);
    });
  }
  return _qrLibPromise;
}
async function startScanner(){document.getElementById('btn-start').style.display='none';document.getElementById('btn-stop').style.display='';const ready=await loadQrLibrary();if(!ready){showToast('Scanner could not load. Check your internet connection.','error');document.getElementById('btn-start').style.display='';document.getElementById('btn-stop').style.display='none';return}html5QrCode=new Html5Qrcode("reader");html5QrCode.start({facingMode:"environment"},{fps:10,qrBox:{width:300,height:180}},onScanSuccess,()=>{}).catch(()=>{showToast('Camera access denied or not available','error');document.getElementById('btn-start').style.display='';document.getElementById('btn-stop').style.display='none'})}
function stopScanner(){if(html5QrCode)html5QrCode.stop().then(()=>{document.getElementById('btn-start').style.display='';document.getElementById('btn-stop').style.display='none'})}
function onScanSuccess(t){const p=t.split(',');if(p.length>1)t=p[1].trim();if(t===lastScan)return;lastScan=t;setTimeout(()=>{lastScan=null},5e3);document.getElementById('scan-display').textContent=t;document.getElementById('scan-edit').value=t;document.getElementById('scan-edit').style.display='none';document.getElementById('scan-display').style.display='block';document.getElementById('btn-edit-scan').textContent='Edit';document.getElementById('scan-confirm').style.display='block';const e=document.getElementById('scan-result');e.textContent='Scan detected — confirm below';e.classList.add('found');playBeep();showToast('Barcode detected — confirm to proceed','info')}
async function confirmScan(){const i=document.getElementById('scan-edit'),d=document.getElementById('scan-display'),v=i.style.display!=='none'?i.value.trim():d.textContent.trim();if(!v){showToast('Empty','error');return}document.getElementById('f-tracking').value=v;document.getElementById('scan-confirm').style.display='none';document.getElementById('scan-result').textContent='Confirmed: '+v;showToast('Tracking: '+v,'success');try{const res=await fetch('/api/check-rescan',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({tracking:v})});const data=await res.json();if(data.ok&&data.is_rescan){showToast('Handoff detected — adding to your list','info');await submitRescan(v)}}catch(e){console.error('Rescan check failed:',e)}}
function toggleEditScan(){const d=document.getElementById('scan-display'),i=document.getElementById('scan-edit'),b=document.getElementById('btn-edit-scan');if(i.style.display==='none'){i.value=d.textContent;d.style.display='none';i.style.display='block';i.focus();i.select();b.textContent='Done'}else{d.textContent=i.value.trim()||'(empty)';i.style.display='none';d.style.display='block';b.textContent='Edit'}}
function rescanBarcode(){document.getElementById('scan-confirm').style.display='none';document.getElementById('scan-result').textContent='Waiting for barcode...';document.getElementById('scan-result').classList.remove('found');lastScan=null;showToast('Ready to scan again','info')}
function playBeep(){try{const c=new(window.AudioContext||window.webkitAudioContext),o=c.createOscillator(),g=c.createGain();o.connect(g);g.connect(c.destination);o.frequency.value=1200;o.type='sine';g.gain.value=.15;o.start();o.stop(c.currentTime+.12)}catch(e){}}
function selectReason(btn){document.querySelectorAll('.reason-btn').forEach(b=>b.classList.remove('selected'));btn.classList.add('selected');selectedReasonCategory=btn.getAttribute('data-category');selectedSubReason='';const wrap=document.getElementById('sub-reason-wrap');if(selectedReasonCategory==='Other'){wrap.style.display='none';selectedSubReason='';return}wrap.style.display='';const subs=REASON_SUBS[selectedReasonCategory]||[];renderDropdownOptions('sub-reason',subs);document.getElementById('f-sub-reason').value='';document.getElementById('sub-reason-placeholder').textContent='-- Select reason --';document.getElementById('sub-reason-placeholder').classList.add('placeholder');document.getElementById('sub-reason-options').querySelectorAll('.custom-dropdown-item').forEach(item=>item.classList.remove('selected'));document.getElementById('custom-reason-wrap').style.display='none';document.getElementById('f-custom-reason').value=''}
function onSubReasonChange(){const v=document.getElementById('f-sub-reason').value;selectedSubReason=v;const cw=document.getElementById('custom-reason-wrap');if(selectedReasonCategory==='Other'){cw.style.display=''}else{cw.style.display='none';document.getElementById('f-custom-reason').value=''}}
async function submitRecord(){const t=document.getElementById('f-tracking').value.trim();if(!t){showToast('Enter a tracking number','error');return}if(!selectedReasonCategory){showToast('Please select a reason','error');return}let subReason=selectedSubReason;if(selectedReasonCategory==='Other'){const notesVal=document.getElementById('f-notes').value.trim();subReason=notesVal||'Other'}else{if(!subReason){showToast('Please select a sub-reason','error');return}}const b=document.getElementById('btn-submit');b.disabled=true;b.textContent='Submitting...';const notes=document.getElementById('f-notes').value.trim();const remarks=selectedReasonCategory==='Other'?'':notes;const r={tracking:t,rider_name:currentRider,hub:currentHub,status:'Pending',notes:remarks,reason_category:selectedReasonCategory,sub_reason:subReason,scanned_at:Date.now(),time:new Date().toLocaleString(),synced:false};try{const resp=await fetch('/api/scan',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(r)});const d=await resp.json();if(d.ok){r.synced=true;showToast('Submitted successfully!','success');stopScanner()}else showToast('Failed: '+(d.error||'Unknown'),'error')}catch(e){showToast('Network error.','error')}history=history.filter(h=>h.tracking!==r.tracking);history.unshift(r);if(history.length>100)history=history.slice(0,100);localStorage.setItem('bt_history',JSON.stringify(history));renderHistory();document.getElementById('f-tracking').value='';document.getElementById('f-notes').value='';document.getElementById('scan-result').textContent='Waiting for barcode...';document.getElementById('scan-result').classList.remove('found');document.getElementById('scan-confirm').style.display='none';document.querySelectorAll('.reason-btn').forEach(b=>b.classList.remove('selected'));document.getElementById('sub-reason-wrap').style.display='none';selectedReasonCategory='';selectedSubReason='';b.disabled=false;b.textContent='Submit'}
/* Handoff rescan: submit without reason selection */
async function submitRescan(tracking){const r={tracking:tracking,rider_name:currentRider,hub:currentHub,status:'Pending',notes:'',reason_category:'',sub_reason:'',scanned_at:Date.now(),time:new Date().toLocaleString(),synced:false};try{const resp=await fetch('/api/scan',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(r)});const d=await resp.json();if(d.ok){r.synced=true;showToast('Package added to your list','success');stopScanner()}else showToast('Failed: '+(d.error||'Unknown'),'error')}catch(e){showToast('Network error.','error')}history=history.filter(h=>h.tracking!==r.tracking);history.unshift(r);if(history.length>100)history=history.slice(0,100);localStorage.setItem('bt_history',JSON.stringify(history));renderHistory();document.getElementById('f-tracking').value='';document.getElementById('scan-result').textContent='Waiting for barcode...';document.getElementById('scan-result').classList.remove('found');document.getElementById('scan-confirm').style.display='none';loadDashboardComments()}
function renderHistory(){const e=document.getElementById('history');if(history.length===0)e.innerHTML='<p style="color:var(--text-muted);text-align:center;padding:20px;font-size:13px">No scans yet</p>';else e.innerHTML=history.map(h=>{const rc=h.reason_category||'';const reasonTag=rc?'<span class="h-status h-reason">'+esc(rc)+'</span>':'';const done=String(h.status||'').toLowerCase()==='completed';const statusTag=done?'<span class="h-status h-done">Completed</span>':'';const sub=h.sub_reason?'<div class="h-sub">'+esc(h.sub_reason)+'</div>':'';return '<div class="history-item"><div class="h-head">'+reasonTag+statusTag+'<span class="h-time">'+esc(h.time||'')+'</span></div><div class="h-tracking">'+esc(h.tracking)+'</div>'+sub+'</div>'}).join('');document.getElementById('s-total').textContent=history.length;document.getElementById('s-synced').textContent=history.filter(h=>h.synced).length;document.getElementById('s-pending').textContent=history.filter(h=>!h.synced).length;const unseenCount=history.filter(h=>isScanUnseen(h)).length;document.getElementById('nav-history-count').textContent=unseenCount;updateBottomBadges()}
function esc(s){const d=document.createElement('div');d.textContent=s;return d.innerHTML}
function showToast(m,t){const e=document.getElementById('toast');e.textContent=m;e.className='toast show '+t;setTimeout(()=>e.className='toast',3e3)}
let _commentRefreshTimer=null;
let _lastCommentRecords=[];
let _lastCommentsHash='';

/* ── Seen-comment state ──────────────────────────────────────────────────
   The green dot means "there is something you have not looked at yet":
   either a brand new parcel, or new feedback on a parcel you already opened.
   The state is kept per rider in localStorage so it survives a reload or a
   re-login - otherwise every login re-lights every old tracking number. */
const SEEN_STORE='bt_seen_comments';
const SEEN_SCANS_STORE='bt_seen_scans';
let _seenState={};
let _seenScansState={};
function dcHash(str){let h=5381;for(let i=0;i<str.length;i++){h=((h<<5)+h+str.charCodeAt(i))|0}return ''+(h>>>0)}
function loadSeenState(rider){
  let all={};
  let allScans={};
  try{all=JSON.parse(localStorage.getItem(SEEN_STORE)||'{}')||{}}catch(e){all={}}
  try{allScans=JSON.parse(localStorage.getItem(SEEN_SCANS_STORE)||'{}')||{}}catch(e){allScans={}}
  _seenState=all;
  _seenScansState=allScans;
}
function saveSeenState(){
  try{localStorage.setItem(SEEN_STORE,JSON.stringify(_seenState))}catch(e){}
}
function saveSeenScansState(){
  try{localStorage.setItem(SEEN_SCANS_STORE,JSON.stringify(_seenScansState))}catch(e){}
}
function seenEntry(tracking){
  const mine=_seenState[currentRider];
  return (mine&&mine[tracking])||null;
}
function markSeen(tracking,row,hash){
  if(!currentRider||!tracking)return;
  if(!_seenState[currentRider])_seenState[currentRider]={};
  _seenState[currentRider][tracking]={r:row,h:hash};
  saveSeenState();
}
/* Check if a scan is unseen (new since last viewed) */
function isScanUnseen(h){
  if(!currentRider)return false;
  const mine=_seenScansState[currentRider];
  if(!mine||!mine[h.tracking])return true;
  return mine[h.tracking]!==h.scanned_at;
}
function markScanSeen(h){
  if(!currentRider||!h.tracking)return;
  if(!_seenScansState[currentRider])_seenScansState[currentRider]={};
  _seenScansState[currentRider][h.tracking]=h.scanned_at;
  saveSeenScansState();
}
function markAllScansSeen(){
  if(!currentRider||!history.length)return;
  if(!_seenScansState[currentRider])_seenScansState[currentRider]={};
  history.forEach(h=>{_seenScansState[currentRider][h.tracking]=h.scanned_at});
  saveSeenScansState();
}
/* Unread = never seen, rescanned (new sheet row), or the comments changed */
function isUnread(r){
  const seen=seenEntry(r.tracking);
  if(!seen)return true;
  return seen.r!==r.row||seen.h!==dcHash(r.comments||'');
}

/* A parcel whose journey was ended is dropped from the rider's Recent Scans */
function pruneEndedScans(records){
  const ended={};
  records.forEach(r=>{if(String(r.status||'').toLowerCase()==='completed')ended[r.tracking]=true});
  const keys=Object.keys(ended);
  if(!keys.length)return;
  const before=history.length;
  history=history.filter(h=>!ended[h.tracking]);
  if(history.length!==before){
    localStorage.setItem('bt_history',JSON.stringify(history));
    renderHistory();
  }
}

async function loadDashboardComments(){
  if(!currentRider)return;
  try{
    const resp=await fetch('/api/my-records?rider='+encodeURIComponent(currentRider));
    const data=await resp.json();
    if(!data.ok)return;
    const records=data.records||[];
    /* Parcels whose journey was ended disappear from this rider's app */
    pruneEndedScans(records);
    const active=records.filter(r=>String(r.status||'Pending').toLowerCase()!=='completed');
    const withComments=active.filter(r=>r.comments&&r.comments.trim());
    /* Build a quick hash to detect changes */
    const hash=withComments.map(r=>r.tracking+'|'+r.comments).join(';;');
    if(hash===_lastCommentsHash)return;
    _lastCommentsHash=hash;
    _lastCommentRecords=withComments;
    renderDashboardComments(withComments);
    const unreadCount=withComments.filter(r=>isUnread(r)).length;
    document.getElementById('nav-comments-count').textContent=unreadCount;
    updateBottomBadges();
  }catch(e){console.error('Failed to load comments:',e)}
}
function renderDashboardComments(records){
  const el=document.getElementById('dash-comments');
  if(records.length===0){el.innerHTML='<div class="dc-empty">No comments from dashboard yet.</div>';return}
  el.innerHTML=records.map(r=>{
    const hasNew=isUnread(r);
    const dotHtml=hasNew?'<span class="dc-new-dot"></span>':'';
    return '<div class="dash-comment-item" onclick="openRiderCommentModal(\''+esc(r.tracking).replace(/'/g,"\\'")+'\')">'+
      dotHtml+
      '<span class="dc-tracking-row">'+esc(r.tracking)+'</span>'+
      '</div>';
  }).join('');
}
let _riderCommentTracking='';
let _riderCommentRefreshTimer=null;

function dcGetAuthorDisplay(author){
  const a=(author||'').trim();
  const lower=a.toLowerCase();
  if(lower==='dashboard'||lower==='staff')return{name:'Staff',type:'staff'};
  if(lower==='seller')return{name:'Seller',type:'seller'};
  if(lower.startsWith('rider'))return{name:a,type:'rider'};
  if(a)return{name:'Rider '+a,type:'rider'};
  return{name:'Staff',type:'staff'};
}
function dcGetInitials(name){
  const parts=name.trim().split(/\s+/);
  if(parts.length>=2)return(parts[0][0]+parts[1][0]).toUpperCase();
  return name.substring(0,2).toUpperCase();
}
function dcFormatTs(ts){
  if(!ts)return'';
  const d=new Date(ts);
  if(isNaN(d.getTime()))return ts;
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  let h=d.getHours();const m=d.getMinutes().toString().padStart(2,'0');
  const ap=h>=12?'p.m.':'a.m.';h=h%12||12;
  return months[d.getMonth()]+' '+d.getDate()+', '+d.getFullYear()+', '+h+':'+m+' '+ap;
}
function renderRiderCommentLine(line){
  const m=line.match(/^\[([^\]]+)\]\s*(.*)$/);
  if(m){
    const ts=m[1];const text=m[2];
    const am=text.match(/^(.*?)\s*\(([^)]+)\)$/);
    if(am){
      const info=dcGetAuthorDisplay(am[2].trim());
      const initials=dcGetInitials(info.name);
      const fts=dcFormatTs(ts);
      return '<div class="dc-comment-entry">'+
        '<div class="dc-comment-avatar '+info.type+'">'+esc(initials)+'</div>'+
        '<div class="dc-comment-content">'+
          '<div class="dc-comment-author-name">'+esc(info.name)+'</div>'+
          '<div class="dc-comment-text">'+esc(am[1].trim())+'</div>'+
          (fts?'<div class="dc-comment-timestamp"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>'+esc(fts)+'</div>':'')+
        '</div></div>';
    }
    const info=dcGetAuthorDisplay('dashboard');
    const initials=dcGetInitials(info.name);
    const fts=dcFormatTs(ts);
    return '<div class="dc-comment-entry">'+
      '<div class="dc-comment-avatar '+info.type+'">'+esc(initials)+'</div>'+
      '<div class="dc-comment-content">'+
        '<div class="dc-comment-author-name">'+esc(info.name)+'</div>'+
        '<div class="dc-comment-text">'+esc(text)+'</div>'+
        (fts?'<div class="dc-comment-timestamp"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>'+esc(fts)+'</div>':'')+
      '</div></div>';
  }
  return '<div class="dc-comment-entry"><div class="dc-comment-content"><div class="dc-comment-text">'+esc(line)+'</div></div></div>';
}

function openRiderCommentModal(tracking){
  _riderCommentTracking=tracking;
  _lastModalCommentsHash='';
  document.getElementById('dc-modal-tracking').textContent=tracking;
  document.getElementById('dc-modal-body').innerHTML='<div class="dc-empty">Loading comments...</div>';
  document.getElementById('dc-modal-input').value='';
  document.getElementById('dc-modal-submit').disabled=false;
  document.getElementById('dc-modal-submit').textContent='Send';
  document.getElementById('dc-modal-end-btn').disabled=false;
  document.getElementById('dc-modal-end-btn').innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg> End';
  const overlay=document.getElementById('dc-modal-overlay');
  overlay.style.display='';
  overlay.classList.add('active');
  setTimeout(()=>document.getElementById('dc-modal-input').focus(),100);
  refreshModalComments();
  if(_riderCommentRefreshTimer)clearInterval(_riderCommentRefreshTimer);
  _riderCommentRefreshTimer=setInterval(refreshModalComments,10000);
  /* Mark as read: the rider is looking at this parcel now */
  if(_lastCommentRecords){
    const rec=_lastCommentRecords.find(r=>r.tracking===tracking);
    if(rec)markSeen(tracking,rec.row,dcHash(rec.comments||''));
    renderDashboardComments(_lastCommentRecords);
  }
  /* Also fetch fresh data as backup */
  try{
    fetch('/api/my-records?rider='+encodeURIComponent(currentRider)).then(r=>r.json()).then(data=>{
      if(!data.ok)return;
      const rec=(data.records||[]).find(r=>r.tracking===tracking);
      if(rec){markSeen(tracking,rec.row,dcHash(rec.comments||''));renderDashboardComments(_lastCommentRecords||[])}
    }).catch(()=>{});
  }catch(e){}
}

function closeRiderCommentModal(){
  const overlay=document.getElementById('dc-modal-overlay');
  overlay.classList.remove('active');
  overlay.style.display='none';
  _riderCommentTracking='';
  if(_riderCommentRefreshTimer){clearInterval(_riderCommentRefreshTimer);_riderCommentRefreshTimer=null}
}

let _lastModalCommentsHash='';
async function refreshModalComments(){
  if(!_riderCommentTracking||!currentRider)return;
  try{
    const resp=await fetch('/api/my-records?rider='+encodeURIComponent(currentRider));
    const data=await resp.json();
    if(!data.ok)return;
    const records=data.records||[];
    const match=records.find(r=>r.tracking===_riderCommentTracking);
    const body=document.getElementById('dc-modal-body');
    if(!match||!match.comments||!match.comments.trim()){
      const emptyHtml='<div class="dc-empty">No comments yet. Add the first one below.</div>';
      if(body.innerHTML!==emptyHtml)body.innerHTML=emptyHtml;
      _lastModalCommentsHash='';
      return;
    }
    const hash=match.comments;
    /* The rider is looking at it right now, so it counts as seen */
    markSeen(_riderCommentTracking,match.row,dcHash(match.comments||''));
    if(hash===_lastModalCommentsHash)return;
    _lastModalCommentsHash=hash;
    const commentLines=match.comments.split('\n').filter(l=>l.trim()).reverse();
    body.innerHTML=commentLines.map(line=>renderRiderCommentLine(line)).join('');
  }catch(e){console.error('Failed to refresh comments:',e)}
}

async function submitRiderCommentFromModal(){
  const input=document.getElementById('dc-modal-input');
  const btn=document.getElementById('dc-modal-submit');
  const val=input.value.trim();
  if(!val||!_riderCommentTracking||!currentRider)return;
  btn.disabled=true;
  btn.textContent='...';
  try{
    const res=await fetch('/api/rider-comment',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rider:currentRider,tracking:_riderCommentTracking,comment:val})});
    const data=await res.json();
    if(data.ok){
      input.value='';
      showToast('Comment sent','success');
      refreshModalComments();
      loadDashboardComments();
    }else{
      showToast('Failed: '+(data.error||'Unknown'),'error');
    }
  }catch(e){showToast('Network error','error')}
  btn.disabled=false;
  btn.textContent='Send';
}

async function endPackageJourney(){
  if(!_riderCommentTracking||!currentRider)return;
  const tracking=_riderCommentTracking;
  const btn=document.getElementById('dc-modal-end-btn');
  btn.disabled=true;
  btn.textContent='...';
  try{
    const res=await fetch('/api/end-package',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rider:currentRider,tracking:tracking})});
    const data=await res.json();
    if(data.ok){
      /* The parcel leaves this rider's app - only a rescan brings it back */
      history=history.filter(h=>h.tracking!==tracking);
      localStorage.setItem('bt_history',JSON.stringify(history));
      showToast('Package journey ended','success');
      renderHistory();
      loadDashboardComments();
      closeRiderCommentModal();
    }else{
      showToast('Failed: '+(data.error||'Unknown'),'error');
    }
  }catch(e){showToast('Failed: '+(e&&e.message?e.message:'Network error'),'error')}
  btn.disabled=false;
  btn.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg> End';
}

document.addEventListener('click',function(e){
  if(e.target.id==='dc-modal-overlay')closeRiderCommentModal();
});
document.addEventListener('keydown',function(e){
  if(e.key==='Escape')closeRiderCommentModal();
});
function startCommentRefresh(){loadDashboardComments();if(_commentRefreshTimer)clearInterval(_commentRefreshTimer);_commentRefreshTimer=setInterval(loadDashboardComments,10000)}
function stopCommentRefresh(){if(_commentRefreshTimer){clearInterval(_commentRefreshTimer);_commentRefreshTimer=null}closeRiderCommentModal()}
loadRiders();
</script>
</body>
</html>"""

# ─── Admin Dashboard HTML (embedded) ─────────────────────────────────────────
ADMIN_HTML = r"""<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Barcode Tracker — Admin</title>
<style>
:root {
  --primary: #F57224; --primary-hover: #E06520; --primary-light: rgba(245,114,36,0.12);
  --success: #22C55E; --success-bg: rgba(34,197,94,0.12);
  --danger: #EF4444; --danger-bg: rgba(239,68,68,0.12);
  --warning: #F59E0B; --info: #3B82F6; --info-bg: rgba(59,130,246,0.12);
  --radius: 12px; --radius-sm: 8px; --radius-xs: 6px;
  --shadow: 0 1px 3px rgba(0,0,0,0.08); --shadow-lg: 0 10px 25px rgba(0,0,0,0.1);
  --transition: 0.2s ease;
  --font: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  --mono: 'JetBrains Mono', 'Fira Code', 'Consolas', monospace;
}
[data-theme="dark"] {
  --bg: #0F0F0F; --bg-card: #1A1A1A; --bg-elevated: #242424; --bg-hover: #2A2A2A; --bg-input: #161922;
  --border: #2A2A2A; --border-light: #333;
  --text: #F5F5F5; --text-secondary: #A3A3A3; --text-muted: #737373;
  --sidebar-bg: #141414; --sidebar-hover: #1F1F1F;
}
[data-theme="light"] {
  --bg: #F8F8F8; --bg-card: #FFF; --bg-elevated: #F5F5F5; --bg-hover: #EEE; --bg-input: #F3F4F6;
  --border: #E5E5E5; --border-light: #D4D4D4;
  --text: #171717; --text-secondary: #525252; --text-muted: #737373;
  --sidebar-bg: #FFF; --sidebar-hover: #F5F5F5;
}
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box}
body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;transition:background var(--transition),color var(--transition)}

/* ── Admin Login ─────────────────────────────────────────────────────── */
#admin-login{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100vh;padding:20px;background:linear-gradient(135deg,var(--bg) 0%,var(--bg-elevated) 100%)}
.admin-login-card{background:var(--bg-card);border-radius:var(--radius);padding:40px 32px;width:100%;max-width:400px;text-align:center;box-shadow:var(--shadow-lg);border:1px solid var(--border)}
.admin-login-card .logo{width:64px;height:64px;background:var(--primary-light);border-radius:16px;display:flex;align-items:center;justify-content:center;margin:0 auto 16px;font-size:28px}
.admin-login-card h1{font-size:22px;font-weight:700;margin-bottom:4px}
.admin-login-card .subtitle{color:var(--text-muted);font-size:13px;margin-bottom:28px}
.admin-login-card .form-group{text-align:left;margin-bottom:16px}
.admin-login-card label{display:block;font-size:12px;font-weight:500;color:var(--text-secondary);margin-bottom:6px}
.admin-login-card input{width:100%;padding:10px 14px;border-radius:var(--radius-sm);border:1px solid var(--border);background:var(--bg-elevated);color:var(--text);font-size:14px;outline:none;transition:border-color var(--transition)}
.admin-login-card input:focus{border-color:var(--primary);box-shadow:0 0 0 3px var(--primary-light)}
.login-error{color:var(--danger);font-size:13px;margin-top:12px;display:none}

/* ── Admin Dashboard Layout ──────────────────────────────────────────── */
#admin-dashboard{display:none}
.app-layout{display:flex;min-height:100vh}
.sidebar{width:240px;background:var(--sidebar-bg);border-right:1px solid var(--border);display:flex;flex-direction:column;position:fixed;top:0;left:0;bottom:0;z-index:100;transition:transform .3s ease}
.sidebar-header{padding:20px;border-bottom:1px solid var(--border)}
.sidebar-header .logo-row{display:flex;align-items:center;gap:12px}
.sidebar-header .logo-icon{width:36px;height:36px;background:var(--primary);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:18px}
.sidebar-header h1{font-size:15px;font-weight:700}
.sidebar-header .admin-badge{display:inline-block;background:var(--primary);color:#fff;font-size:10px;padding:2px 8px;border-radius:8px;font-weight:600;margin-left:8px;vertical-align:middle}
.sidebar-header .refresh-info{font-size:11px;color:var(--text-muted);margin-top:8px}
.sidebar-header .refresh-info b{color:var(--success);font-weight:600}
.sidebar-nav{flex:1;padding:12px;overflow-y:auto}
.nav-item{display:flex;align-items:center;gap:12px;padding:10px 14px;border-radius:var(--radius-sm);color:var(--text-secondary);font-size:13px;font-weight:500;cursor:pointer;transition:all var(--transition);margin-bottom:2px}
.nav-item:hover{background:var(--sidebar-hover);color:var(--text)}
.nav-item.active{background:var(--primary-light);color:var(--primary);font-weight:600}
.nav-item .icon{font-size:16px;width:20px;text-align:center}
.nav-item .badge{margin-left:auto;background:var(--primary);color:#fff;font-size:10px;padding:2px 7px;border-radius:10px;font-weight:600}
.sidebar-footer{padding:12px;border-top:1px solid var(--border)}
.sidebar-footer .stats-mini{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:10px}
.stat-mini{text-align:center;padding:8px 4px;background:var(--bg-elevated);border-radius:var(--radius-xs);font-size:10px;color:var(--text-muted)}
.stat-mini b{display:block;font-size:16px;color:var(--text);font-weight:700}
.theme-toggle{width:100%;padding:8px 12px;background:var(--bg-elevated);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text-secondary);font-size:12px;cursor:pointer;display:flex;align-items:center;gap:8px;transition:all var(--transition)}
.theme-toggle:hover{border-color:var(--primary);color:var(--primary)}
.btn-logout{width:100%;padding:8px;background:var(--bg-elevated);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text-secondary);font-size:12px;cursor:pointer;transition:all var(--transition);margin-top:8px}
.btn-logout:hover{border-color:var(--danger);color:var(--danger)}

/* ── Main Content ────────────────────────────────────────────────────── */
.main-content{flex:1;margin-left:240px;padding:24px;max-width:1200px}
.section{display:none}
.section.active{display:block}
.section-title{font-size:18px;font-weight:700;margin-bottom:20px;display:flex;align-items:center;gap:10px}
.section-title .icon{font-size:20px}

/* ── Stats Cards ─────────────────────────────────────────────────────── */
.stats-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:24px}
.stat-card{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:18px;text-align:center;transition:border-color var(--transition)}
.stat-card:hover{border-color:var(--border-light)}
.stat-card .stat-value{font-size:28px;font-weight:700;color:var(--text)}
.stat-card .stat-label{font-size:11px;color:var(--text-muted);margin-top:4px;text-transform:uppercase;letter-spacing:.5px}
.stat-card.accent-orange .stat-value{color:var(--primary)}
.stat-card.accent-green .stat-value{color:var(--success)}
.stat-card.accent-blue .stat-value{color:var(--info)}
.stat-card.accent-red .stat-value{color:var(--danger)}

/* ── Filters ─────────────────────────────────────────────────────────── */
.filters-bar{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:20px}
.filters-bar input,.filters-bar select{padding:9px 14px;border-radius:var(--radius-sm);border:1px solid var(--border);background:var(--bg-card);color:var(--text);font-size:13px;outline:none;transition:border-color var(--transition);min-width:140px}
.filters-bar input:focus,.filters-bar select:focus{border-color:var(--primary);box-shadow:0 0 0 3px var(--primary-light)}
.filters-bar input[type="text"]{flex:1;min-width:200px}
.btn-clear{background:transparent;border:1px solid var(--border);color:var(--text-secondary);padding:9px 18px;border-radius:var(--radius-sm);font-size:13px;cursor:pointer;transition:all var(--transition)}
.btn-clear:hover{border-color:var(--primary);color:var(--primary)}

/* ── Data Table ──────────────────────────────────────────────────────── */
.table-container{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden}
.data-table{width:100%;border-collapse:collapse;font-size:13px}
.data-table th{background:var(--bg-elevated);color:var(--text-muted);font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;padding:12px 14px;text-align:left;border-bottom:1px solid var(--border);position:sticky;top:0;z-index:1}
.data-table td{padding:12px 14px;border-bottom:1px solid var(--border);vertical-align:middle}
.data-table tr:last-child td{border-bottom:none}
.data-table tr:hover{background:var(--bg-hover)}
.data-table .tracking-cell{font-family:var(--mono);font-size:12px;color:var(--text);word-break:break-all}
.data-table .rider-cell{color:var(--text);font-weight:500}
.data-table .hub-cell{color:var(--text-muted);font-size:12px}
.data-table .time-cell{color:var(--text-muted);font-size:12px;white-space:nowrap}
.data-table .notes-cell{color:var(--text-muted);font-size:12px;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}

/* ── Status badges ───────────────────────────────────────────────────── */
.status-badge{display:inline-block;padding:4px 12px;border-radius:12px;font-size:11px;font-weight:600}
.status-badge.pending{background:var(--primary-light);color:var(--primary)}
.status-badge.in-transit{background:var(--info-bg);color:var(--info)}
.status-badge.delivered{background:var(--success-bg);color:var(--success)}

/* ── Status dropdown ─────────────────────────────────────────────────── */
.status-select{padding:6px 10px;border-radius:var(--radius-xs);border:1px solid var(--border);background:var(--bg-elevated);color:var(--text);font-size:12px;outline:none;cursor:pointer;transition:border-color var(--transition)}
.status-select:focus{border-color:var(--primary)}

/* ── Empty state ─────────────────────────────────────────────────────── */
.empty-state{text-align:center;padding:60px 20px;color:var(--text-muted)}
.empty-state .icon{font-size:48px;margin-bottom:12px;opacity:.5}
.empty-state h3{color:var(--text);font-size:16px;margin-bottom:6px}
.empty-state p{font-size:13px}

/* ── Loading ─────────────────────────────────────────────────────────── */
.loading-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.6);z-index:999;align-items:center;justify-content:center}
.loading-overlay.show{display:flex}
.loading-spinner{color:#fff;font-size:16px;background:var(--bg-card);padding:20px 32px;border-radius:var(--radius);border:1px solid var(--border)}

/* ── Toast ───────────────────────────────────────────────────────────── */
.toast{position:fixed;bottom:24px;right:24px;padding:14px 20px;border-radius:var(--radius-sm);font-size:13px;font-weight:500;z-index:10001;opacity:0;transform:translateY(20px);transition:all .3s ease;pointer-events:none;box-shadow:var(--shadow-lg)}
.toast.show{opacity:1;transform:translateY(0)}
.toast.success{background:var(--success);color:#fff}
.toast.error{background:var(--danger);color:#fff}
.toast.info{background:var(--primary);color:#fff}

/* ── WhatsApp Section ────────────────────────────────────────────────── */
.wa-card{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:20px;margin-bottom:20px}
.wa-card .card-title{font-size:12px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:14px}
.wa-connection{display:flex;align-items:center;gap:12px;padding:14px;background:var(--bg-elevated);border-radius:var(--radius-sm);margin-bottom:16px}
.wa-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.wa-dot.connected{background:var(--success);box-shadow:0 0 8px rgba(34,197,94,.4)}
.wa-dot.disconnected{background:var(--danger)}
.wa-conn-label{font-size:13px;font-weight:500;flex:1}
.wa-group-row{display:flex;gap:10px;align-items:center}
.wa-group-row input{flex:1;padding:10px 14px;border-radius:var(--radius-sm);border:1px solid var(--border);background:var(--bg-elevated);color:var(--text);font-size:13px;outline:none;transition:border-color var(--transition)}
.wa-group-row input:focus{border-color:var(--success);box-shadow:0 0 0 3px rgba(34,197,94,.15)}
.btn-wa-check{background:transparent;border:1px solid var(--success);color:var(--success);padding:10px 18px;border-radius:var(--radius-sm);font-size:13px;font-weight:500;cursor:pointer;transition:all var(--transition);white-space:nowrap}
.btn-wa-check:hover{background:var(--success-bg)}
.btn-wa-send{background:var(--success);color:#fff;border:none;padding:6px 14px;border-radius:var(--radius-xs);font-size:12px;font-weight:600;cursor:pointer;transition:all var(--transition);white-space:nowrap}
.btn-wa-send:hover{background:#16A34A}
.btn-wa-send:disabled{opacity:.5;cursor:not-allowed}
.btn-wa-send.sent{background:var(--success)}

/* ── Mobile ──────────────────────────────────────────────────────────── */
.mobile-header{display:none;position:fixed;top:0;left:0;right:0;height:56px;background:var(--sidebar-bg);border-bottom:1px solid var(--border);padding:0 16px;align-items:center;gap:12px;z-index:99}
.hamburger{width:36px;height:36px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:5px;cursor:pointer;border-radius:var(--radius-xs)}
.hamburger:hover{background:var(--bg-hover)}
.hamburger span{width:18px;height:2px;background:var(--text);border-radius:2px}
.mobile-header h1{font-size:15px;font-weight:700}
.sidebar-overlay{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);z-index:99}
@media(max-width:768px){
  .mobile-header{display:flex}
  .sidebar{transform:translateX(-100%)}
  .sidebar.open{transform:translateX(0)}
  .sidebar-overlay.show{display:block}
  .main-content{margin-left:0;padding:72px 16px 16px}
  .stats-row{grid-template-columns:repeat(2,1fr)}
  .filters-bar{flex-direction:column}
  .filters-bar input,.filters-bar select{width:100%}
  .table-container{overflow-x:auto}
}
</style>
</head>
<body>

<!-- ═══════════════════ ADMIN LOGIN ═══════════════════ -->
<div id="admin-login">
  <div class="admin-login-card">
    <div class="logo">&#128230;</div>
    <h1>Admin Dashboard</h1>
    <p class="subtitle">Barcode Tracker — Management Panel</p>
    <div class="form-group">
      <label>Password</label>
      <input type="password" id="admin-password" placeholder="Enter admin password" onkeypress="if(event.key==='Enter')adminLogin()">
    </div>
    <button class="btn btn-primary" style="width:100%;padding:10px;background:var(--primary);color:#fff;border:none;border-radius:var(--radius-sm);font-size:14px;font-weight:600;cursor:pointer" onclick="adminLogin()">Login</button>
    <div id="admin-login-error" class="login-error"></div>
  </div>
</div>

<!-- ═══════════════════ ADMIN DASHBOARD ═══════════════════ -->
<div id="admin-dashboard">
<div class="mobile-header"><div class="hamburger" onclick="toggleSidebar()"><span></span><span></span><span></span></div><h1>Admin Panel</h1></div>
<div class="sidebar-overlay" id="sidebar-overlay" onclick="toggleSidebar()"></div>
<div class="app-layout">
  <aside class="sidebar" id="sidebar">
    <div class="sidebar-header">
      <div class="logo-row"><div class="logo-icon">&#128230;</div><h1>Admin<span class="admin-badge">ADMIN</span></h1></div>
      <div class="refresh-info">Auto-refresh: <b id="refresh-countdown">10</b>s</div>
    </div>
    <nav class="sidebar-nav">
      <div class="nav-item active" onclick="switchSection('dashboard',this)"><span class="icon">&#128202;</span> Dashboard</div>
      <div class="nav-item" onclick="switchSection('records',this)"><span class="icon">&#128203;</span> Records <span class="badge" id="nav-record-count">0</span></div>
      <div class="nav-item" onclick="switchSection('whatsapp',this)"><span class="icon">&#128172;</span> WhatsApp</div>
    </nav>
    <div class="sidebar-footer">
      <div class="stats-mini">
        <div class="stat-mini"><b id="s-total">0</b>Total</div>
        <div class="stat-mini"><b id="s-today">0</b>Today</div>
        <div class="stat-mini"><b id="s-pending">0</b>Pending</div>
        <div class="stat-mini"><b id="s-delivered">0</b>Delivered</div>
      </div>
      <button class="theme-toggle" onclick="toggleTheme()"><span id="theme-icon">&#9728;</span> <span id="theme-label">Dark Mode</span></button>
      <button class="btn-logout" onclick="adminLogout()">Logout</button>
    </div>
  </aside>
  <main class="main-content">

    <!-- ── Dashboard Section ─────────────────────────────────────── -->
    <div class="section active" id="sec-dashboard">
      <div class="section-title"><span class="icon">&#128202;</span> Overview</div>
      <div class="stats-row">
        <div class="stat-card accent-orange"><div class="stat-value" id="stat-total">0</div><div class="stat-label">Total Scans</div></div>
        <div class="stat-card"><div class="stat-value" id="stat-today">0</div><div class="stat-label">Today</div></div>
        <div class="stat-card accent-red"><div class="stat-value" id="stat-pending">0</div><div class="stat-label">Pending</div></div>
        <div class="stat-card accent-blue"><div class="stat-value" id="stat-transit">0</div><div class="stat-label">In Transit</div></div>
        <div class="stat-card accent-green"><div class="stat-value" id="stat-delivered">0</div><div class="stat-label">Delivered</div></div>
        <div class="stat-card"><div class="stat-value" id="stat-riders">0</div><div class="stat-label">Active Riders</div></div>
      </div>
      <div class="section-title" style="font-size:14px;margin-top:8px"><span class="icon">&#128337;</span> Recent Activity</div>
      <div class="table-container" style="max-height:400px;overflow-y:auto">
        <table class="data-table">
          <thead><tr><th>Tracking</th><th>Rider</th><th>Hub</th><th>Scanned</th><th>Status</th></tr></thead>
          <tbody id="recent-table-body"></tbody>
        </table>
      </div>
    </div>

    <!-- ── Records Section ───────────────────────────────────────── -->
    <div class="section" id="sec-records">
      <div class="section-title"><span class="icon">&#128203;</span> All Records</div>
      <div class="filters-bar">
        <input type="text" id="filter-search" placeholder="Search tracking number..." oninput="applyFilters()">
        <select id="filter-hub" onchange="applyFilters()"><option value="">All Hubs</option></select>
        <select id="filter-rider" onchange="applyFilters()"><option value="">All Riders</option></select>
        <select id="filter-status" onchange="applyFilters()">
          <option value="">All Statuses</option>
          <option value="Pending">Pending</option>
          <option value="In Transit">In Transit</option>
          <option value="Delivered">Delivered</option>
        </select>
        <input type="date" id="filter-date" onchange="applyFilters()">
        <button class="btn-clear" onclick="clearFilters()">Clear</button>
      </div>
      <div class="table-container" style="max-height:600px;overflow-y:auto">
        <table class="data-table">
          <thead><tr><th>#</th><th>Tracking</th><th>Rider</th><th>Hub</th><th>Scanned At</th><th>Status</th><th>Reason</th><th>Sub-Reason</th><th>Notes</th><th>Action</th><th>WhatsApp</th></tr></thead>
          <tbody id="table-body"></tbody>
        </table>
        <div id="empty-state" class="empty-state" style="display:none">
          <div class="icon">&#128230;</div>
          <h3>No records found</h3>
          <p>Try adjusting your filters or wait for new scans.</p>
        </div>
      </div>
    </div>

    <!-- ── WhatsApp Section ──────────────────────────────────────── -->
    <div class="section" id="sec-whatsapp">
      <div class="section-title"><span class="icon">&#128172;</span> WhatsApp</div>
      <div class="wa-card">
        <div class="card-title">Connection</div>
        <div class="wa-connection">
          <span id="wa-status" class="wa-dot disconnected"></span>
          <span class="wa-conn-label" id="wa-conn-label">Not connected</span>
          <button class="btn-wa-check" onclick="checkWhatsAppStatus()">Check</button>
        </div>
        <div class="wa-group-row">
          <input type="text" id="wa-group-name" placeholder="Enter WhatsApp group name...">
        </div>
      </div>
      <div class="wa-card">
        <div class="card-title">Send Notes to Group</div>
        <p style="color:var(--text-muted);font-size:13px;margin-bottom:16px">Go to Records tab and click "Send Notes" button on any record with notes. The message will be sent to the group above.</p>
        <div id="wa-records-list"></div>
      </div>
    </div>

  </main>
</div>
</div>

<!-- Loading Overlay -->
<div class="loading-overlay" id="loading-overlay">
  <div class="loading-spinner">Updating status...</div>
</div>

<div class="toast" id="toast"></div>

<script>
let allRecords=[],filteredRecords=[],refreshTimer=null,countdown=10;

function toggleTheme(){const h=document.documentElement,c=h.getAttribute('data-theme'),n=c==='dark'?'light':'dark';h.setAttribute('data-theme',n);document.getElementById('theme-icon').innerHTML=n==='dark'?'&#9790;':'&#9728;';document.getElementById('theme-label').textContent=n==='dark'?'Light Mode':'Dark Mode'}
function toggleSidebar(){document.getElementById('sidebar').classList.toggle('open');document.getElementById('sidebar-overlay').classList.toggle('show')}
function switchSection(id,el){document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));document.getElementById('sec-'+id).classList.add('active');document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));if(el)el.classList.add('active');if(window.innerWidth<=768)toggleSidebar()}

// ── Admin Login ──────────────────────────────────────────────────────────
async function adminLogin(){
  const pw=document.getElementById('admin-password').value;
  if(!pw){showLoginError('Enter password');return}
  try{
    const resp=await fetch('/api/admin/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:pw})});
    const data=await resp.json();
    if(data.ok){document.getElementById('admin-login').style.display='none';document.getElementById('admin-dashboard').style.display='block';startAutoRefresh()}
    else showLoginError(data.error||'Wrong password');
  }catch(e){showLoginError('Network error')}
}
function showLoginError(msg){const el=document.getElementById('admin-login-error');el.textContent=msg;el.style.display=''}
function adminLogout(){if(refreshTimer)clearInterval(refreshTimer);document.getElementById('admin-dashboard').style.display='none';document.getElementById('admin-login').style.display='';document.getElementById('admin-password').value=''}

// ── Auto Refresh ─────────────────────────────────────────────────────────
function startAutoRefresh(){loadRecords();refreshTimer=setInterval(()=>{countdown--;document.getElementById('refresh-countdown').textContent=countdown;if(countdown<=0){loadRecords();countdown=10}},1000)}
async function loadRecords(){
  try{const resp=await fetch('/api/admin/records');const data=await resp.json();if(data.ok){allRecords=data.records||[];populateFilterOptions();applyFilters();updateStats()}}
  catch(e){console.error('Failed to load records:',e)}
}

// ── Stats ────────────────────────────────────────────────────────────────
function updateStats(){
  const today=new Date().toISOString().split('T')[0];
  const todayRecords=allRecords.filter(r=>r.scanned_at&&r.scanned_at.startsWith(today));
  const riders=new Set(allRecords.map(r=>r.rider).filter(Boolean));
  document.getElementById('stat-total').textContent=allRecords.length;
  document.getElementById('stat-today').textContent=todayRecords.length;
  document.getElementById('stat-pending').textContent=allRecords.filter(r=>r.status==='Pending').length;
  document.getElementById('stat-transit').textContent=allRecords.filter(r=>r.status==='In Transit').length;
  document.getElementById('stat-delivered').textContent=allRecords.filter(r=>r.status==='Delivered').length;
  document.getElementById('stat-riders').textContent=riders.size;
  document.getElementById('s-total').textContent=allRecords.length;
  document.getElementById('s-today').textContent=todayRecords.length;
  document.getElementById('s-pending').textContent=allRecords.filter(r=>r.status==='Pending').length;
  document.getElementById('s-delivered').textContent=allRecords.filter(r=>r.status==='Delivered').length;
  document.getElementById('nav-record-count').textContent=allRecords.length;
  renderRecentActivity();
}

// ── Recent Activity (Dashboard) ──────────────────────────────────────────
function renderRecentActivity(){
  const tbody=document.getElementById('recent-table-body');
  const recent=allRecords.slice(-10).reverse();
  if(recent.length===0){tbody.innerHTML='<tr><td colspan="5" style="text-align:center;color:var(--text-muted);padding:30px">No records yet</td></tr>';return}
  tbody.innerHTML=recent.map(r=>`<tr><td class="tracking-cell">${esc(r.tracking)}</td><td class="rider-cell">${esc(r.rider)}</td><td class="hub-cell">${esc(r.hub)}</td><td class="time-cell">${esc(r.scanned_at)}</td><td><span class="status-badge ${statusClass(r.status)}">${esc(r.status)}</span></td></tr>`).join('');
}

// ── Filters ──────────────────────────────────────────────────────────────
function populateFilterOptions(){
  const hubs=[...new Set(allRecords.map(r=>r.hub).filter(Boolean))].sort();
  const riders=[...new Set(allRecords.map(r=>r.rider).filter(Boolean))].sort();
  const hubSel=document.getElementById('filter-hub'),curH=hubSel.value;
  hubSel.innerHTML='<option value="">All Hubs</option>';hubs.forEach(h=>{const o=document.createElement('option');o.value=h;o.textContent=h;hubSel.appendChild(o)});hubSel.value=curH;
  const riderSel=document.getElementById('filter-rider'),curR=riderSel.value;
  riderSel.innerHTML='<option value="">All Riders</option>';riders.forEach(r=>{const o=document.createElement('option');o.value=r;o.textContent=r;riderSel.appendChild(o)});riderSel.value=curR;
}
function applyFilters(){
  const search=document.getElementById('filter-search').value.toLowerCase(),hub=document.getElementById('filter-hub').value,rider=document.getElementById('filter-rider').value,status=document.getElementById('filter-status').value,date=document.getElementById('filter-date').value;
  filteredRecords=allRecords.filter(r=>{
    if(search&&!r.tracking.toLowerCase().includes(search))return false;
    if(hub&&r.hub!==hub)return false;if(rider&&r.rider!==rider)return false;
    if(status&&r.status!==status)return false;
    if(date&&r.scanned_at&&!r.scanned_at.startsWith(date))return false;
    return true;
  });
  renderTable();
}
function clearFilters(){document.getElementById('filter-search').value='';document.getElementById('filter-hub').value='';document.getElementById('filter-rider').value='';document.getElementById('filter-status').value='';document.getElementById('filter-date').value='';applyFilters()}

// ── Table Rendering ──────────────────────────────────────────────────────
function renderTable(){
  const tbody=document.getElementById('table-body'),empty=document.getElementById('empty-state');
  if(filteredRecords.length===0){tbody.innerHTML='';empty.style.display='';return}
  empty.style.display='none';
  tbody.innerHTML=filteredRecords.map((r,i)=>`<tr>
    <td>${r.row}</td>
    <td class="tracking-cell">${esc(r.tracking)}</td>
    <td class="rider-cell">${esc(r.rider)}</td>
    <td class="hub-cell">${esc(r.hub)}</td>
    <td class="time-cell">${esc(r.scanned_at)}</td>
    <td><span class="status-badge ${statusClass(r.status)}">${esc(r.status)}</span></td>
    <td>${r.reason_category?'<span class="status-badge pending">'+esc(r.reason_category)+'</span>':'<span style="color:var(--text-muted)">-</span>'}</td>
    <td style="color:var(--text-secondary);font-size:12px">${esc(r.sub_reason)||'-'}</td>
    <td class="notes-cell" title="${esc(r.notes)}">${esc(r.notes)}</td>
    <td><select class="status-select" onchange="updateStatus(${r.row},this.value)"><option value="Pending" ${r.status==='Pending'?'selected':''}>Pending</option><option value="In Transit" ${r.status==='In Transit'?'selected':''}>In Transit</option><option value="Delivered" ${r.status==='Delivered'?'selected':''}>Delivered</option></select></td>
    <td>${(r.notes||r.reason_category)?'<button class="btn-wa-send" id="wa-btn-'+r.row+'" onclick="sendToWhatsApp('+r.row+')">Send</button>':'<span style="color:var(--text-muted)">-</span>'}</td>
  </tr>`).join('');
}
function statusClass(s){if(s==='Pending')return'pending';if(s==='In Transit')return'in-transit';if(s==='Delivered')return'delivered';return''}
function esc(s){if(!s)return'';const d=document.createElement('div');d.textContent=s;return d.innerHTML}

// ── Status Update ────────────────────────────────────────────────────────
async function updateStatus(row,newStatus){
  document.getElementById('loading-overlay').classList.add('show');
  try{
    const resp=await fetch('/api/admin/update-status',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({row,status:newStatus})});
    const data=await resp.json();
    if(data.ok){showToast('Status updated to '+newStatus,'success');const rec=allRecords.find(r=>r.row===row);if(rec)rec.status=newStatus;updateStats();renderTable()}
    else showToast('Failed: '+(data.error||'Unknown'),'error');
  }catch(e){showToast('Network error','error')}
  document.getElementById('loading-overlay').classList.remove('show');
}

// ── Toast ────────────────────────────────────────────────────────────────
function showToast(msg,type){const t=document.getElementById('toast');t.textContent=msg;t.className='toast show '+type;setTimeout(()=>t.className='toast',3000)}

// ── WhatsApp ─────────────────────────────────────────────────────────────
async function checkWhatsAppStatus(){
  try{
    const resp=await fetch('/api/admin/whatsapp-status');const data=await resp.json();
    const dot=document.getElementById('wa-status'),label=document.getElementById('wa-conn-label');
    if(data.ok&&data.logged_in){dot.className='wa-dot connected';label.textContent='Connected to WhatsApp Web';showToast('WhatsApp Web is connected','success')}
    else{dot.className='wa-dot disconnected';label.textContent='Not connected';showToast('WhatsApp not connected. Run: python whatsapp_sender.py setup','error')}
  }catch(e){showToast('Could not check WhatsApp status','error')}
}
function getRemarks(rec){
  let notes=(rec.notes||'').trim();
  if(!notes)return '';
  const cat=(rec.reason_category||'').trim();
  const sub=(rec.sub_reason||'').trim();
  const prefix=cat+(cat&&sub?' > '+sub:'');
  if(prefix&&notes.indexOf(prefix)===0){
    notes=notes.slice(prefix.length).trim().replace(/^[>\-\s|:]+/,'').trim();
    const p=notes.indexOf('|');
    if(p>=0)notes=notes.slice(p+1).trim();
  }
  if(!notes)return '';
  if(notes===cat||notes===sub)return '';
  if(cat&&sub&&notes===cat+' > '+sub)return '';
  return notes;
}
async function sendToWhatsApp(row){
  const rec=allRecords.find(r=>r.row===row);if(!rec){showToast('Record not found','error');return}
  const groupName=document.getElementById('wa-group-name').value.trim();
  if(!groupName){showToast('Enter a WhatsApp group name first','error');document.getElementById('wa-group-name').focus();return}
  if(!rec.notes&&!rec.reason_category){showToast('No notes or reason to send','error');return}
  const btn=document.getElementById('wa-btn-'+row);
  if(btn){btn.disabled=true;btn.textContent='...'}
  const remarks=getRemarks(rec);
  const message='*Barcode Tracker Update*\nTracking: '+rec.tracking+'\nRider: '+rec.rider+'\nHub: '+rec.hub+'\nStatus: '+rec.status+'\nScanned: '+rec.scanned_at+(rec.reason_category?'\nReason: '+rec.reason_category+(rec.sub_reason?' > '+rec.sub_reason:''):'')+(remarks?'\nNotes: '+remarks:'');
  try{
    const resp=await fetch('/api/admin/send-whatsapp',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({group_name:groupName,message:message})});
    const data=await resp.json();
    if(data.success){showToast('Sent to '+groupName,'success');if(btn){btn.textContent='Sent!';btn.className='btn-wa-send sent'}}
    else{showToast('Failed: '+(data.error||'Unknown'),'error');if(btn){btn.disabled=false;btn.textContent='Send'}}
  }catch(e){showToast('Network error','error');if(btn){btn.disabled=false;btn.textContent='Send'}}
}
</script>
</body>
</html>"""

# ─── Main ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Only open browser for local development
    if os.environ.get('PORT') is None:
        url = f"http://localhost:{PORT}"
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
        print(f"\n  Barcode Tracker running at {url}\n  Press Ctrl+C to stop.\n")
    else:
        print(f"\n  Barcode Tracker running on port {PORT}\n")
    app.run(host="0.0.0.0", port=PORT, debug=False)
