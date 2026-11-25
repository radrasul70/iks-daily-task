# -*- coding: utf-8 -*-
import http.server
import socketserver
import webbrowser
import socket
import json
import os
import io
import ssl
import subprocess
import threading
import sqlite3
import secrets
import hashlib
import hmac
import time
import http.cookies
import shutil
from contextlib import closing

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.chart import BarChart, Reference
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

PORT = int(os.environ.get("PORT", 8000))
HTTPS_PORT = int(os.environ.get("HTTPS_PORT", 8443))
IS_CLOUD = bool(os.environ.get("RENDER") or os.environ.get("RAILWAY_ENVIRONMENT") or os.environ.get("DYNO"))

import sys
if getattr(sys, "frozen", False):
    DIRECTORY = os.path.dirname(os.path.abspath(sys.executable))
else:
    DIRECTORY = os.path.dirname(os.path.abspath(__file__))
# On Render, durable data must live under the attached persistent disk mount.
# Locally, keep the existing v18 folder layout unchanged.
DATA_DIR = os.environ.get("DATA_DIR", DIRECTORY)
os.makedirs(DATA_DIR, exist_ok=True)
DB_FILE = os.path.join(DATA_DIR, "tasks_db.sqlite3")
LEGACY_DB_FILE = os.path.join(DATA_DIR, "tasks_db_clean.json")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
SESSION_TTL_SECONDS = 12 * 60 * 60
SESSIONS = {}
SESSION_LOCK = threading.Lock()
PBKDF2_ITERATIONS = 310000
CERT_FILE = os.path.join(DIRECTORY, "cert.pem")
KEY_FILE = os.path.join(DIRECTORY, "key.pem")
DB_LOCK = threading.Lock()

DEPARTMENTS = [
    "کارشناس تولیدی",
    "کارشناس غیر تولیدی",
    "امور قراردادها",
    "کلیه کارشناس ها"
]

CATEGORIES = [
    "فعالیت اجرایی",
    "پیگیری",
    "اهداف عملیاتی",
    "اهداف و برنامه‌ها",
    "تسک"
]

PRIORITIES = [
    "بحرانی",
    "بالا",
    "متوسط",
    "پایین"
]

STATUSES = [
    "شروع نشده",
    "در حال انجام",
    "در انتظار تایید",
    "متوقف شده",
    "لغو شده",
    "تکمیل شده"
]

# Local test login remains admin/123 for compatibility. Cloud deployments
# must provide a private initial password through the INITIAL_ADMIN_PASSWORD
# environment variable before the first database is created.
INITIAL_ADMIN_PASSWORD = os.environ.get("INITIAL_ADMIN_PASSWORD", "123")
if IS_CLOUD and len(INITIAL_ADMIN_PASSWORD) < 12:
    raise RuntimeError("Set INITIAL_ADMIN_PASSWORD to a unique password of at least 12 characters in the hosting environment.")

DEFAULT_USERS = [
    {
        "username": "admin",
        "password": INITIAL_ADMIN_PASSWORD,
        "name": "مدیر بازرگانی",
        "role": "مدیر بازرگانی",
        "department": "مدیریت بازرگانی"
    }
]

SEED_TASKS = []

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def map_legacy_category(cat):
    if not cat:
        return "فعالیت اجرایی"
    if cat in CATEGORIES:
        return cat
    if "پیگیری" in cat:
        return "پیگیری"
    if "عملیاتی" in cat:
        return "اهداف عملیاتی"
    if "برنامه" in cat:
        return "اهداف و برنامه‌ها"
    if "تسک" in cat:
        return "تسک"
    return "فعالیت اجرایی"

def normalize_task(t):
    done = bool(t.get("done", False))
    status = t.get("status") or ("تکمیل شده" if done else "در حال انجام")
    if status == "تکمیل شده":
        done = True
    progress = int(t.get("progress", 100 if done else 0) or 0)
    if done and progress < 100:
        progress = 100
    return {
        "id": int(t.get("id", 1)),
        "category": map_legacy_category(t.get("category")),
        "title": t.get("title") or "",
        "department": t.get("department") or "کارشناس تولیدی",
        "priority": t.get("priority") or "متوسط",
        "status": status,
        "assigneeUsername": t.get("assigneeUsername") or "",
        "assigneeName": t.get("assigneeName") or "نامشخص",
        "month": t.get("month") or "",
        "startDate": t.get("startDate") or "",
        "endDate": t.get("endDate") or "",
        "completionDate": t.get("completionDate") or "",
        "progress": progress,
        "alarm": t.get("alarm") or "",
        "result": t.get("result") or "",
        "nextAction": t.get("nextAction") or "",
        "notes": t.get("notes") or "",
        "createdByUsername": t.get("createdByUsername") or "",
        "createdByName": t.get("createdByName") or "",
        "createdByRole": t.get("createdByRole") or "",
        # Preserve manager referral marker across database save/load cycles.
        "referredByManager": bool(t.get("referredByManager", False)),
        "done": done,
        "alerted": bool(t.get("alerted", False)),
        "updatedAt": int(t.get("updatedAt", 0) or 0)
    }

def db_connect():
    conn = sqlite3.connect(DB_FILE, timeout=30, isolation_level=None)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=30000")
    return conn

def hash_password(password):
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", str(password).encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return "pbkdf2_sha256${}${}${}".format(PBKDF2_ITERATIONS, salt.hex(), digest.hex())

def verify_password(password, stored):
    if not stored:
        return False
    if not str(stored).startswith("pbkdf2_sha256$"):
        return hmac.compare_digest(str(password), str(stored))
    try:
        _, iterations, salt_hex, digest_hex = str(stored).split("$", 3)
        actual = hashlib.pbkdf2_hmac("sha256", str(password).encode("utf-8"), bytes.fromhex(salt_hex), int(iterations)).hex()
        return hmac.compare_digest(actual, digest_hex)
    except Exception:
        return False

def public_user(profile):
    item = dict(profile or {})
    item.pop("password_hash", None)
    item["password"] = ""
    return item

def init_database():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    with closing(db_connect()) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("""CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            profile_json TEXT NOT NULL,
            password_hash TEXT NOT NULL
        )""")
        conn.execute("""CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY,
            payload_json TEXT NOT NULL,
            assignee_username TEXT NOT NULL DEFAULT '',
            created_by_username TEXT NOT NULL DEFAULT '',
            department TEXT NOT NULL DEFAULT ''
        )""")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_tasks_assignee ON tasks(assignee_username)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_tasks_creator ON tasks(created_by_username)")
        conn.execute("CREATE TABLE IF NOT EXISTS task_tombstones (task_id INTEGER PRIMARY KEY, deleted_at INTEGER NOT NULL, deleted_by TEXT NOT NULL DEFAULT '')")
        conn.execute("CREATE TABLE IF NOT EXISTS app_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
        user_count = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        if user_count == 0:
            legacy = None
            if os.path.exists(LEGACY_DB_FILE):
                try:
                    with open(LEGACY_DB_FILE, "r", encoding="utf-8") as f:
                        legacy = json.load(f)
                    backup_path = LEGACY_DB_FILE + ".pre_sqlite_backup"
                    if not os.path.exists(backup_path):
                        shutil.copy2(LEGACY_DB_FILE, backup_path)
                except Exception:
                    legacy = None
            users = legacy.get("users", []) if isinstance(legacy, dict) else []
            if not isinstance(users, list) or not users:
                users = DEFAULT_USERS
            for user in users:
                profile = dict(user)
                username = str(profile.get("username", "")).strip()
                password = str(profile.pop("password", "123"))
                if not username:
                    continue
                conn.execute("INSERT OR REPLACE INTO users(username, profile_json, password_hash) VALUES(?,?,?)",
                             (username, json.dumps(profile, ensure_ascii=False), hash_password(password)))
            tasks = legacy.get("tasks", []) if isinstance(legacy, dict) else []
            for task in tasks:
                normalized = normalize_task(task)
                conn.execute("INSERT OR REPLACE INTO tasks(id,payload_json,assignee_username,created_by_username,department) VALUES(?,?,?,?,?)",
                    (normalized["id"], json.dumps(normalized, ensure_ascii=False), normalized.get("assigneeUsername", ""), normalized.get("createdByUsername", ""), normalized.get("department", "")))
            next_num = legacy.get("nextNumber", 1) if isinstance(legacy, dict) else 1
            conn.execute("INSERT OR REPLACE INTO app_meta(key,value) VALUES('nextNumber',?)", (str(max(int(next_num or 1), 1)),))
        if conn.execute("SELECT 1 FROM app_meta WHERE key='nextNumber'").fetchone() is None:
            max_id = conn.execute("SELECT COALESCE(MAX(id),0) FROM tasks").fetchone()[0]
            conn.execute("INSERT INTO app_meta(key,value) VALUES('nextNumber',?)", (str(max_id + 1),))

def load_db():
    init_database_if_needed()
    with closing(db_connect()) as conn:
        tasks = [json.loads(row[0]) for row in conn.execute("SELECT payload_json FROM tasks ORDER BY id")]
        users = []
        for row in conn.execute("SELECT profile_json FROM users ORDER BY username"):
            users.append(public_user(json.loads(row[0])))
        meta = conn.execute("SELECT value FROM app_meta WHERE key='nextNumber'").fetchone()
        next_num = int(meta[0]) if meta else (max([t.get("id", 0) for t in tasks], default=0) + 1)
        return {"tasks": tasks, "nextNumber": next_num, "users": users}

_DB_INITIALIZED = False
_DB_INIT_LOCK = threading.Lock()
def init_database_if_needed():
    global _DB_INITIALIZED
    if _DB_INITIALIZED:
        return
    with _DB_INIT_LOCK:
        if not _DB_INITIALIZED:
            init_database()
            _DB_INITIALIZED = True

def save_users_from_payload(users):
    if not isinstance(users, list) or not users:
        raise ValueError("لیست کاربران نامعتبر است.")
    with closing(db_connect()) as conn:
        conn.execute("BEGIN IMMEDIATE")
        incoming_names = set()
        for raw_user in users:
            profile = dict(raw_user)
            username = str(profile.get("username", "")).strip()
            if not username:
                continue
            incoming_names.add(username)
            new_password = str(profile.pop("password", "") or "")
            existing = conn.execute("SELECT password_hash FROM users WHERE username=?", (username,)).fetchone()
            if existing:
                pwd_hash = hash_password(new_password) if new_password else existing[0]
            else:
                if not new_password:
                    raise ValueError("برای کاربر جدید تعیین رمز عبور الزامی است.")
                pwd_hash = hash_password(new_password)
            conn.execute("INSERT OR REPLACE INTO users(username,profile_json,password_hash) VALUES(?,?,?)",
                         (username, json.dumps(profile, ensure_ascii=False), pwd_hash))
        if not incoming_names:
            raise ValueError("حداقل یک کاربر باید باقی بماند.")
        existing_names = {r[0] for r in conn.execute("SELECT username FROM users")}
        for removed in existing_names - incoming_names:
            conn.execute("DELETE FROM users WHERE username=?", (removed,))
        conn.execute("COMMIT")

def get_user_record(username):
    with closing(db_connect()) as conn:
        row = conn.execute("SELECT profile_json,password_hash FROM users WHERE username=?", (username,)).fetchone()
        if not row:
            return None
        return json.loads(row[0]), row[1]

def set_user_password(username, password):
    with closing(db_connect()) as conn:
        cur = conn.execute("UPDATE users SET password_hash=? WHERE username=?", (hash_password(password), username))
        return cur.rowcount > 0

def create_session(username):
    token = secrets.token_urlsafe(32)
    with SESSION_LOCK:
        SESSIONS[token] = {"username": username, "expires": time.time() + SESSION_TTL_SECONDS}
    return token

def session_username(handler):
    cookie = http.cookies.SimpleCookie()
    try:
        cookie.load(handler.headers.get("Cookie", ""))
        morsel = cookie.get("session")
        token = morsel.value if morsel else ""
    except Exception:
        token = ""
    if not token:
        return None
    with SESSION_LOCK:
        record = SESSIONS.get(token)
        if not record:
            return None
        if record["expires"] < time.time():
            SESSIONS.pop(token, None)
            return None
        record["expires"] = time.time() + SESSION_TTL_SECONDS
        return record["username"]

def current_user(handler):
    username = session_username(handler)
    if not username:
        return None
    record = get_user_record(username)
    if not record:
        return None
    return public_user(record[0])

def is_manager_user(user):
    return bool(user and user.get("role") == "مدیر بازرگانی")

def task_visible_to_user(task, user):
    """Experts see only tasks assigned to them (legacy self-created unassigned tasks included)."""
    if is_manager_user(user):
        return True
    username = user.get("username", "")
    return (task.get("assigneeUsername") == username or
            (not task.get("assigneeUsername") and task.get("createdByUsername") == username))

def save_tasks_for_user(incoming, next_number, user, deleted_ids=None):
    if not isinstance(incoming, list):
        raise ValueError("فهرست تسک‌ها نامعتبر است.")
    manager = is_manager_user(user)
    username = user.get("username", "")
    id_map = {}
    with closing(db_connect()) as conn:
        conn.execute("BEGIN IMMEDIATE")
        all_rows = conn.execute("SELECT id,payload_json FROM tasks").fetchall()
        old_by_id = {r[0]: json.loads(r[1]) for r in all_rows}
        max_id = max(old_by_id.keys(), default=0)
        for raw in incoming:
            task = normalize_task(raw)
            requested_id = task["id"]
            old = old_by_id.get(requested_id)
            tombstone = conn.execute("SELECT deleted_at FROM task_tombstones WHERE task_id=?", (requested_id,)).fetchone()
            incoming_stamp = int(task.get("updatedAt", 0) or 0)
            if tombstone and incoming_stamp <= int(tombstone[0] or 0):
                continue
            if tombstone and incoming_stamp > int(tombstone[0] or 0):
                conn.execute("DELETE FROM task_tombstones WHERE task_id=?", (requested_id,))
            if not manager:
                if old and not task_visible_to_user(old, user):
                    # A stale copy of a task reassigned by the manager must never
                    # recreate it or modify it. Only a genuinely new local task
                    # (created by this user) may receive a replacement global ID.
                    if task.get("createdByUsername") == username and old.get("createdByUsername") != username:
                        max_id += 1
                        task["id"] = max_id
                        id_map[str(requested_id)] = max_id
                        old = None
                    else:
                        continue
                if old:
                    task["assigneeUsername"] = old.get("assigneeUsername") or username
                    task["assigneeName"] = old.get("assigneeName") or user.get("name", "")
                    task["createdByUsername"] = old.get("createdByUsername") or username
                    task["createdByName"] = old.get("createdByName") or user.get("name", "")
                    task["createdByRole"] = old.get("createdByRole", "")
                    task["referredByManager"] = old.get("referredByManager", False)
                else:
                    task["assigneeUsername"] = username
                    task["assigneeName"] = user.get("name", "")
                    task["createdByUsername"] = username
                    task["createdByName"] = user.get("name", "")
                    task["createdByRole"] = user.get("role", "")
                    task["referredByManager"] = False
            # Last-write-wins by per-task update timestamp. A timestamp is only
            # compared for the same task; stale offline copies cannot overwrite newer data.
            if old:
                try:
                    incoming_stamp = int(task.get("updatedAt", 0) or 0)
                    server_stamp = int(old.get("updatedAt", 0) or 0)
                except Exception:
                    incoming_stamp, server_stamp = 0, 0
                if incoming_stamp < server_stamp:
                    continue
            conn.execute("INSERT OR REPLACE INTO tasks(id,payload_json,assignee_username,created_by_username,department) VALUES(?,?,?,?,?)",
                (task["id"], json.dumps(task, ensure_ascii=False), task.get("assigneeUsername", ""), task.get("createdByUsername", ""), task.get("department", "")))
            old_by_id[task["id"]] = task

        # Explicit deletions are restricted to the manager or the task creator.
        for raw_delete in (deleted_ids or []):
            try:
                if isinstance(raw_delete, dict):
                    task_id = int(raw_delete.get("id"))
                    deleted_at = int(raw_delete.get("deletedAt", 0) or 0)
                else:
                    task_id = int(raw_delete)
                    deleted_at = int(time.time() * 1000)
            except Exception:
                continue
            old = old_by_id.get(task_id)
            if not old:
                continue
            if manager or old.get("createdByUsername") == username:
                if deleted_at >= int(old.get("updatedAt", 0) or 0):
                    conn.execute("DELETE FROM tasks WHERE id=?", (task_id,))
                    conn.execute("INSERT OR REPLACE INTO task_tombstones(task_id,deleted_at,deleted_by) VALUES(?,?,?)", (task_id, deleted_at, username))

        max_id = conn.execute("SELECT COALESCE(MAX(id),0) FROM tasks").fetchone()[0]
        final_next = max(int(next_number or 1), max_id + 1)
        conn.execute("INSERT OR REPLACE INTO app_meta(key,value) VALUES('nextNumber',?)", (str(final_next),))
        conn.execute("COMMIT")
    return {"idMap": id_map, "nextNumber": final_next}

def get_visible_state(user):
    db = load_db()
    db["tasks"] = [t for t in db["tasks"] if task_visible_to_user(t, user)]
    if not is_manager_user(user):
        db["users"] = [u for u in db["users"] if u.get("username") == user.get("username")]
    return db

def backup_database():
    init_database_if_needed()
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    destination = os.path.join(BACKUP_DIR, "tasks_db_{}.sqlite3".format(stamp))
    with closing(db_connect()) as source:
        with closing(sqlite3.connect(destination)) as target:
            source.backup(target)
    # Keep the latest 30 daily snapshots; avoid unbounded disk growth.
    files = sorted([os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR) if f.endswith(".sqlite3")], key=os.path.getmtime, reverse=True)
    for old_path in files[30:]:
        try:
            os.remove(old_path)
        except OSError:
            pass
    return destination

def backup_scheduler():
    last_day = None
    while True:
        today = time.strftime("%Y%m%d")
        if today != last_day:
            try:
                backup_database()
                last_day = today
                print("نسخه پشتیبان روزانه دیتابیس ایجاد شد.")
            except Exception as exc:
                print("خطا در تهیه نسخه پشتیبان:", exc)
        time.sleep(300)

def ensure_self_signed_cert(local_ip):
    if os.path.exists(CERT_FILE) and os.path.exists(KEY_FILE):
        return True
    try:
        subprocess.run([
            "openssl", "req", "-x509", "-newkey", "rsa:2048",
            "-keyout", KEY_FILE, "-out", CERT_FILE,
            "-days", "365", "-nodes",
            "-subj", f"/CN={local_ip}"
        ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False

def build_corporate_xlsx(tasks, users):
    if not HAS_OPENPYXL:
        return None
    wb = openpyxl.Workbook()
    ws_all = wb.active
    ws_all.title = "All"
    ws_all.views.sheetView[0].rightToLeft = True

    headers = [
        "شماره (ID)",
        "دسته‌بندی (Category)",
        "شرح اقدام (Description)",
        "دپارتمان (Department)",
        "اولویت (Priority)",
        "وضعیت (Status)",
        "مسئول (Responsible)",
        "ماه (Month)",
        "تاریخ شروع (Start Date)",
        "تاریخ پایان (End Date)",
        "تاریخ تکمیل (Completion Date)",
        "درصد پیشرفت (Progress %)",
        "آلارم یادآور (Alarm)",
        "نتیجه (Result)",
        "اقدام بعدی (Next Action)",
        "توضیحات (Notes)"
    ]

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Tahoma", size=10, bold=True, color="FFFFFF")
    cell_font = Font(name="Tahoma", size=10)
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    done_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    cat_fills = {
        "فعالیت اجرایی": PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid"),
        "پیگیری": PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid"),
        "اهداف عملیاتی": PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
        "اهداف و برنامه‌ها": PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid"),
        "تسک": PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"),
    }
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws_all.append(headers)
    ws_all.row_dimensions[1].height = 28
    for col_idx in range(1, len(headers) + 1):
        c = ws_all.cell(1, col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

    sorted_tasks = sorted(tasks, key=lambda x: x.get("id", 0))
    for r_idx, t in enumerate(sorted_tasks, start=2):
        prog_val = (float(t.get("progress", 0)) / 100.0)
        row_data = [
            t.get("id", r_idx - 1),
            t.get("category", ""),
            t.get("title", ""),
            t.get("department", ""),
            t.get("priority", ""),
            t.get("status", ""),
            t.get("assigneeName", ""),
            t.get("month", ""),
            t.get("startDate", ""),
            t.get("endDate", ""),
            t.get("completionDate", ""),
            prog_val,
            t.get("alarm", ""),
            t.get("result", ""),
            t.get("nextAction", ""),
            t.get("notes", "")
        ]
        ws_all.append(row_data)
        ws_all.row_dimensions[r_idx].height = 24
        is_done = (t.get("status") == "تکمیل شده" or t.get("done"))
        row_fill = done_fill if is_done else (zebra_fill if r_idx % 2 == 1 else PatternFill(fill_type=None))
        for c_idx in range(1, len(headers) + 1):
            cell = ws_all.cell(r_idx, c_idx)
            cell.font = cell_font
            cell.border = border
            if c_idx == 2 and cell.value in cat_fills:
                cell.fill = cat_fills[cell.value]
            elif row_fill.fill_type:
                cell.fill = row_fill
            if c_idx == 12:
                cell.number_format = "0%"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in (1, 2, 8, 9, 10, 11, 13):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

    col_widths = [12, 22, 46, 20, 14, 16, 18, 12, 15, 15, 16, 16, 20, 26, 26, 28]
    for i, w in enumerate(col_widths, start=1):
        ws_all.column_dimensions[get_column_letter(i)].width = w

    ws_all.freeze_panes = "A2"
    ws_all.auto_filter.ref = f"A1:P{max(2, len(sorted_tasks) + 1)}"

    max_dv_row = max(200, len(sorted_tasks) + 50)
    dv_cat = DataValidation(type="list", formula1='"' + ",".join(CATEGORIES) + '"', allow_blank=True)
    dv_dept = DataValidation(type="list", formula1='"' + ",".join(DEPARTMENTS) + '"', allow_blank=True)
    dv_prio = DataValidation(type="list", formula1='"' + ",".join(PRIORITIES) + '"', allow_blank=True)
    dv_stat = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
    ws_all.add_data_validation(dv_cat)
    ws_all.add_data_validation(dv_dept)
    ws_all.add_data_validation(dv_prio)
    ws_all.add_data_validation(dv_stat)
    dv_cat.add(f"B2:B{max_dv_row}")
    dv_dept.add(f"D2:D{max_dv_row}")
    dv_prio.add(f"E2:E{max_dv_row}")
    dv_stat.add(f"F2:F{max_dv_row}")

    def add_summary_sheet(sheet_name, title_label, col_letter_in_all, items_list, add_chart=True):
        ws = wb.create_sheet(title=sheet_name)
        ws.views.sheetView[0].rightToLeft = True
        ws["A1"] = title_label
        ws["A1"].font = Font(name="Tahoma", size=12, bold=True, color="1E3A8A")
        ws["A3"] = "عنوان"
        ws["B3"] = "تعداد کل اقدامات"
        ws["C3"] = "تکمیل شده"
        ws["D3"] = "میانگین پیشرفت"
        for col_i in range(1, 5):
            hc = ws.cell(3, col_i)
            hc.fill = header_fill
            hc.font = header_font
            hc.alignment = Alignment(horizontal="center", vertical="center")
            hc.border = border

        for idx, item_val in enumerate(items_list, start=4):
            ws.cell(idx, 1, item_val).font = cell_font
            ws.cell(idx, 1).border = border
            c_cnt = ws.cell(idx, 2, f'=COUNTIF(All!${col_letter_in_all}$2:${col_letter_in_all}$500, A{idx})')
            c_done = ws.cell(idx, 3, f'=COUNTIFS(All!${col_letter_in_all}$2:${col_letter_in_all}$500, A{idx}, All!$F$2:$F$500, "تکمیل شده")')
            c_avg = ws.cell(idx, 4, f'=IFERROR(AVERAGEIF(All!${col_letter_in_all}$2:${col_letter_in_all}$500, A{idx}, All!$L$2:$L$500), 0)')
            for c_cell in (c_cnt, c_done, c_avg):
                c_cell.font = cell_font
                c_cell.border = border
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
            c_avg.number_format = "0%"

        tot_row = 4 + len(items_list)
        ws.cell(tot_row, 1, "جمع کل").font = Font(name="Tahoma", size=10, bold=True)
        ws.cell(tot_row, 2, f"=SUM(B4:B{tot_row-1})").font = Font(name="Tahoma", size=10, bold=True)
        ws.cell(tot_row, 3, f"=SUM(C4:C{tot_row-1})").font = Font(name="Tahoma", size=10, bold=True)
        ws.cell(tot_row, 4, "=IFERROR(AVERAGE(All!$L$2:$L$500), 0)").font = Font(name="Tahoma", size=10, bold=True)
        ws.cell(tot_row, 4).number_format = "0%"
        for col_i in range(1, 5):
            ws.cell(tot_row, col_i).border = border
            ws.cell(tot_row, col_i).fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 18

        if add_chart and len(items_list) > 0:
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = title_label
            chart.y_axis.title = "تعداد"
            data_ref = Reference(ws, min_col=2, max_col=3, min_row=3, max_row=tot_row - 1)
            cats_ref = Reference(ws, min_col=1, min_row=4, max_row=tot_row - 1)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            chart.width = 16
            chart.height = 10
            ws.add_chart(chart, "F3")

    responsible_names = sorted(list({u.get("name") for u in users if u.get("name")} | {t.get("assigneeName") for t in sorted_tasks if t.get("assigneeName")}))
    months_list = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"]

    add_summary_sheet("department", "گزارش اقدامات به تفکیک دپارتمان", "D", DEPARTMENTS)
    add_summary_sheet("status", "گزارش اقدامات به تفکیک وضعیت", "F", STATUSES)
    add_summary_sheet("priority", "گزارش اقدامات به تفکیک اولویت", "E", PRIORITIES)
    add_summary_sheet("responsible", "گزارش اقدامات به تفکیک مسئولین", "G", responsible_names)
    add_summary_sheet("category", "گزارش اقدامات به تفکیک دسته‌بندی", "B", CATEGORIES)
    add_summary_sheet("درصد پیشرفت ماهانه", "گزارش درصد پیشرفت ماهانه", "H", months_list)

    ws_dash = wb.create_sheet(title="dashboard")
    ws_dash.views.sheetView[0].rightToLeft = True
    ws_dash["A1"] = "داشبورد مدیریتی پیگیری اقدامات روزانه بازرگانی"
    ws_dash["A1"].font = Font(name="Tahoma", size=14, bold=True, color="1E3A8A")
    ws_dash["A3"] = "کل اقدامات ثبت شده"
    ws_dash["B3"] = "=COUNTA(All!C2:C500)"
    ws_dash["A4"] = "اقدامات تکمیل شده"
    ws_dash["B4"] = '=COUNTIF(All!F2:F500, "تکمیل شده")'
    ws_dash["A5"] = "اقدامات در حال انجام"
    ws_dash["B5"] = '=COUNTIF(All!F2:F500, "در حال انجام")'
    ws_dash["A6"] = "میانگین کل پیشرفت"
    ws_dash["B6"] = "=IFERROR(AVERAGE(All!L2:L500), 0)"
    ws_dash["B6"].number_format = "0.0%"
    for r_i in range(3, 7):
        ws_dash.cell(r_i, 1).font = Font(name="Tahoma", size=10, bold=True)
        ws_dash.cell(r_i, 2).font = Font(name="Tahoma", size=11, bold=True, color="0F766E")
        ws_dash.cell(r_i, 1).border = border
        ws_dash.cell(r_i, 2).border = border
    ws_dash.column_dimensions["A"].width = 26
    ws_dash.column_dimensions["B"].width = 16

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "same-origin")
        super().end_headers()

    def send_json(self, code, obj, extra_headers=None):
        payload = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        for name, value in (extra_headers or []):
            self.send_header(name, value)
        self.end_headers()
        self.wfile.write(payload)

    def read_json_body(self):
        length = min(int(self.headers.get("Content-Length", 0)), 5_000_000)
        raw = self.rfile.read(length).decode("utf-8")
        return json.loads(raw) if raw else {}

    def origin_is_same(self):
        origin = self.headers.get("Origin")
        if not origin:
            return True
        from urllib.parse import urlparse
        try:
            return urlparse(origin).netloc.lower() == self.headers.get("Host", "").lower()
        except Exception:
            return False

    def require_user(self):
        user = current_user(self)
        if not user:
            self.send_json(401, {"ok": False, "error": "نشست شما معتبر نیست؛ لطفاً دوباره وارد شوید."})
            return None
        return user

    def do_GET(self):
        if self.path == "/healthz":
            self.send_json(200, {"ok": True, "service": "daily-task"})
            return
        if self.path == "/" or self.path == "":
            self.path = "/index.html"
        if self.path == "/api/auth/me":
            user = current_user(self)
            if user:
                self.send_json(200, {"ok": True, "user": user})
            else:
                self.send_json(401, {"ok": False})
            return
        if self.path.startswith("/api/"):
            user = self.require_user()
            if not user:
                return
            if self.path.startswith("/api/state"):
                self.send_json(200, get_visible_state(user))
                return
            if self.path.startswith("/api/export.xlsx"):
                db = get_visible_state(user)
                xlsx_bytes = build_corporate_xlsx(db.get("tasks", []), db.get("users", []))
                if not xlsx_bytes:
                    self.send_json(501, {"ok": False, "error": "کتابخانه openpyxl نصب نیست."})
                    return
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", 'attachment; filename="IKS-DailyTask-Tracking.xlsx"')
                self.send_header("Content-Length", str(len(xlsx_bytes)))
                self.end_headers()
                self.wfile.write(xlsx_bytes)
                return
            self.send_json(404, {"ok": False, "error": "مسیر پیدا نشد."})
            return
        from urllib.parse import urlsplit, unquote
        static_path = unquote(urlsplit(self.path).path).lstrip("/")
        allowed_static = {"index.html", "logo-imen.jpg", "icon.svg", "manifest.json", "service-worker.js"}
        if static_path in allowed_static:
            self.path = "/" + static_path
            super().do_GET()
        else:
            self.send_error(404, "Not Found")

    def do_POST(self):
        if not self.origin_is_same():
            self.send_json(403, {"ok": False, "error": "درخواست با مبدأ نامعتبر رد شد."})
            return
        try:
            body = self.read_json_body()
        except Exception as e:
            self.send_json(400, {"ok": False, "error": "بدنه درخواست نامعتبر است: " + str(e)})
            return

        if self.path == "/api/login":
            uname = str(body.get("username", "")).strip()
            pwd = str(body.get("password", ""))
            record = get_user_record(uname)
            if record and verify_password(pwd, record[1]):
                # Upgrade legacy password hashes after a successful login.
                if not str(record[1]).startswith("pbkdf2_sha256$"):
                    set_user_password(uname, pwd)
                token = create_session(uname)
                forwarded_proto = str(self.headers.get("X-Forwarded-Proto", "")).lower()
                secure = "; Secure" if (isinstance(self.connection, ssl.SSLSocket) or forwarded_proto == "https") else ""
                cookie = "session={}; HttpOnly; SameSite=Lax; Path=/; Max-Age={}{}".format(token, SESSION_TTL_SECONDS, secure)
                self.send_json(200, {"ok": True, "user": public_user(record[0])}, [("Set-Cookie", cookie)])
                return
            self.send_json(401, {"ok": False, "error": "نام کاربری یا رمز عبور اشتباه است."})
            return

        if self.path == "/api/logout":
            cookie = http.cookies.SimpleCookie()
            cookie.load(self.headers.get("Cookie", ""))
            morsel = cookie.get("session")
            if morsel:
                with SESSION_LOCK:
                    SESSIONS.pop(morsel.value, None)
            self.send_json(200, {"ok": True}, [("Set-Cookie", "session=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0")])
            return

        user = self.require_user()
        if not user:
            return

        if self.path == "/api/change-password":
            old_password = str(body.get("currentPassword", ""))
            new_password = str(body.get("newPassword", ""))
            if len(new_password) < 8:
                self.send_json(400, {"ok": False, "error": "رمز جدید باید حداقل ۸ نویسه داشته باشد."})
                return
            record = get_user_record(user["username"])
            if not record or not verify_password(old_password, record[1]):
                self.send_json(400, {"ok": False, "error": "رمز عبور فعلی صحیح نیست."})
                return
            set_user_password(user["username"], new_password)
            self.send_json(200, {"ok": True})
            return

        if self.path == "/api/users":
            if not is_manager_user(user):
                self.send_json(403, {"ok": False, "error": "فقط مدیر بازرگانی به مدیریت کاربران دسترسی دارد."})
                return
            try:
                incoming_users = body.get("users")
                if not isinstance(incoming_users, list) or not any(str(x.get("username", "")) == user.get("username") for x in incoming_users if isinstance(x, dict)):
                    self.send_json(400, {"ok": False, "error": "حساب مدیر واردشده نمی‌تواند از فهرست کاربران حذف شود."})
                    return
                save_users_from_payload(incoming_users)
                db = load_db()
                self.send_json(200, {"ok": True, "users": db["users"]})
            except Exception as e:
                self.send_json(400, {"ok": False, "error": str(e)})
            return

        if self.path == "/api/tasks":
            try:
                result = save_tasks_for_user(body.get("tasks", []), body.get("nextNumber", 1), user, body.get("deletedIds", []))
                state = get_visible_state(user)
                self.send_json(200, {"ok": True, **result, "tasks": state.get("tasks", []), "users": state.get("users", [])})
            except Exception as e:
                self.send_json(400, {"ok": False, "error": str(e)})
            return

        self.send_json(404, {"ok": False, "error": "مسیر پیدا نشد."})


class ReusableTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    allow_reuse_address = True
    daemon_threads = True

def main():
    os.chdir(DIRECTORY)
    init_database_if_needed()
    threading.Thread(target=backup_scheduler, daemon=True).start()

    if IS_CLOUD:
        print(f"Cloud server running on port {PORT}...")
        with ReusableTCPServer(("0.0.0.0", PORT), Handler) as httpd:
            httpd.serve_forever()
        return

    local_ip = get_local_ip()
    pc_url = f"http://localhost:{PORT}/index.html"
    phone_https = f"https://{local_ip}:{HTTPS_PORT}/index.html"
    has_https = ensure_self_signed_cert(local_ip)

    print("=" * 74)
    print("  سامانه دیلی‌تسک بازرگانی - نسخه نمایشی ۱۹ (مبنای ۱۸.۱) فعال شد!")
    print(f"  1) آدرس در کامپیوتر: {pc_url}")
    if has_https:
        print(f"  2) آدرس امن در گوشی (روی همان Wi-Fi): {phone_https}")
    print("  اطلاعات قبلی در صورت وجود به‌صورت خودکار به SQLite منتقل می‌شود.")
    print("=" * 74)

    webbrowser.open(pc_url)

    if has_https:
        https_server = ReusableTCPServer(("0.0.0.0", HTTPS_PORT), Handler)
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.load_cert_chain(certfile=CERT_FILE, keyfile=KEY_FILE)
        https_server.socket = ctx.wrap_socket(https_server.socket, server_side=True)
        t = threading.Thread(target=https_server.serve_forever, daemon=True)
        t.start()

    with ReusableTCPServer(("0.0.0.0", PORT), Handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nبرنامه متوقف شد.")

if __name__ == "__main__":
    main()
