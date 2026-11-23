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

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(DIRECTORY, "tasks_db.json")
CERT_FILE = os.path.join(DIRECTORY, "cert.pem")
KEY_FILE = os.path.join(DIRECTORY, "key.pem")
DB_LOCK = threading.Lock()

DEPARTMENTS = [
    "کارشناس تولیدی",
    "کارشناس غیر تولیدی",
    "کارشناس امور قراردادها"
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

DEFAULT_USERS = [
    {
        "username": "admin",
        "password": "123",
        "name": "مدیر بازرگانی",
        "role": "مدیر بازرگانی",
        "department": "مدیریت بازرگانی"
    }
]

SEED_TASKS = []

SEED_TASKS = [
    {
        "id": 1,
        "category": "پیگیری",
        "title": "انعقاد قرارداد پارچه تدبیر تکس کوتد و انکوتد",
        "department": "کارشناس تولیدی",
        "priority": "متوسط",
        "status": "تکمیل شده",
        "assigneeUsername": "mansouri",
        "assigneeName": "منصوری",
        "month": "تیر",
        "startDate": "1405/02/08",
        "endDate": "1405/02/15",
        "completionDate": "1405/02/20",
        "progress": 100,
        "alarm": "",
        "result": "تعیین تکلیف نهایی",
        "nextAction": "-",
        "notes": "عدم انعقاد قرارداد",
        "createdByUsername": "admin",
        "createdByName": "مدیر بازرگانی",
        "createdByRole": "مدیر بازرگانی",
        "done": True,
        "alerted": False
    },
    {
        "id": 2,
        "category": "فعالیت اجرایی",
        "title": "خرید نخ داخلی از شرکت شمیم - نمونه اولیه - تیراژ انبوه",
        "department": "کارشناس تولیدی",
        "priority": "بالا",
        "status": "تکمیل شده",
        "assigneeUsername": "tafaghodi",
        "assigneeName": "تفقدی",
        "month": "تیر",
        "startDate": "1405/04/20",
        "endDate": "1405/05/04",
        "completionDate": "1405/05/10",
        "progress": 100,
        "alarm": "",
        "result": "تامین انجام شد",
        "nextAction": "تحویل به انبار",
        "notes": "",
        "createdByUsername": "admin",
        "createdByName": "مدیر بازرگانی",
        "createdByRole": "مدیر بازرگانی",
        "done": True,
        "alerted": False
    },
    {
        "id": 3,
        "category": "اهداف و برنامه‌ها",
        "title": "تمدید بازه قراردادهای باز سال جدید",
        "department": "امور قراردادها",
        "priority": "متوسط",
        "status": "تکمیل شده",
        "assigneeUsername": "jamali",
        "assigneeName": "جمالی",
        "month": "تیر",
        "startDate": "1405/05/01",
        "endDate": "1405/05/08",
        "completionDate": "1405/05/08",
        "progress": 100,
        "alarm": "",
        "result": "الحاقیه تمدید امضا شد",
        "nextAction": "ارسال به مالی",
        "notes": "",
        "createdByUsername": "admin",
        "createdByName": "مدیر بازرگانی",
        "createdByRole": "مدیر بازرگانی",
        "done": True,
        "alerted": False
    },
    {
        "id": 4,
        "category": "اهداف عملیاتی",
        "title": "نحوه تسویه و تعیین تکلیف ضایعات فلزی پیمانکاران پرس کاری (لیست قیمت جدید - اجرای پایلوت - صورتجلسه)",
        "department": "کارشناس غیر تولیدی",
        "priority": "بحرانی",
        "status": "در حال انجام",
        "assigneeUsername": "niroomand",
        "assigneeName": "نیرومند",
        "month": "مرداد",
        "startDate": "1405/05/10",
        "endDate": "1405/05/29",
        "completionDate": "",
        "progress": 50,
        "alarm": "",
        "result": "در حال استعلام نرخ ضایعات",
        "nextAction": "برگزاری جلسه نهایی با پیمانکار",
        "notes": "اولویت ویژه مدیریت",
        "createdByUsername": "admin",
        "createdByName": "مدیر بازرگانی",
        "createdByRole": "مدیر بازرگانی",
        "done": False,
        "alerted": False
    }
]

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
    if cat == "امور قراردادها":
        return "کارشناس امور قراردادها"
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
        "done": done,
        "alerted": bool(t.get("alerted", False))
    }

def load_db():
    with DB_LOCK:
        if not os.path.exists(DB_FILE):
            initial = {
                "tasks": SEED_TASKS,
                "nextNumber": len(SEED_TASKS) + 1,
                "users": DEFAULT_USERS
            }
            with open(DB_FILE, "w", encoding="utf-8") as f:
                json.dump(initial, f, ensure_ascii=False, indent=2)
            return initial
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    users = data.get("users")
                    if not isinstance(users, list) or len(users) == 0:
                        users = DEFAULT_USERS
                    tasks = [normalize_task(x) for x in data.get("tasks", [])]
                    max_id = max([x["id"] for x in tasks], default=0)
                    next_num = max(int(data.get("nextNumber", 1)), max_id + 1)
                    return {
                        "tasks": tasks,
                        "nextNumber": next_num,
                        "users": users
                    }
        except Exception:
            pass
        return {"tasks": SEED_TASKS, "nextNumber": 1, "users": DEFAULT_USERS}

def save_db(data):
    with DB_LOCK:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

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
        super().end_headers()

    def send_json(self, code, obj):
        payload = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        if self.path == "/" or self.path == "":
            self.path = "/index.html"
        if self.path.startswith("/api/state"):
            data = load_db()
            self.send_json(200, data)
            return
        if self.path.startswith("/api/export.xlsx"):
            db = load_db()
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
        super().do_GET()

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length).decode("utf-8")
        try:
            body = json.loads(raw) if raw else {}
        except Exception as e:
            self.send_json(400, {"ok": False, "error": str(e)})
            return

        if self.path.startswith("/api/login"):
            uname = str(body.get("username", "")).strip()
            pwd = str(body.get("password", "")).strip()
            db = load_db()
            users = db.get("users") if isinstance(db.get("users"), list) and len(db.get("users")) > 0 else DEFAULT_USERS
            for u in users:
                if str(u.get("username", "")).strip() == uname and str(u.get("password", "")).strip() == pwd:
                    self.send_json(200, {"ok": True, "user": u})
                    return
            self.send_json(401, {"ok": False, "error": "نام کاربری یا رمز عبور اشتباه است."})
            return

        if self.path.startswith("/api/users"):
            db = load_db()
            users = body.get("users")
            if isinstance(users, list) and len(users) > 0:
                db["users"] = users
                save_db(db)
                self.send_json(200, {"ok": True, "users": db["users"]})
            else:
                self.send_json(400, {"ok": False, "error": "لیست کاربران نامعتبر است."})
            return

        if self.path.startswith("/api/tasks"):
            db = load_db()
            incoming = body.get("tasks", db.get("tasks", []))
            db["tasks"] = [normalize_task(x) for x in incoming]
            db["nextNumber"] = body.get("nextNumber", db.get("nextNumber", 1))
            save_db(db)
            self.send_json(200, {"ok": True})
            return

        self.send_response(404)
        self.end_headers()

class ReusableTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    allow_reuse_address = True
    daemon_threads = True

def main():
    os.chdir(DIRECTORY)
    load_db()

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
    print("  برنامه دیلی تسک صوتی بازرگانی (نسخه نهایی - جامع ۱۶ ستونی + آلارم + داشبورد) فعال شد!")
    print(f"  1) آدرس در کامپیوتر: {pc_url}")
    if has_https:
        print(f"  2) آدرس امن در گوشی (روی همان Wi-Fi): {phone_https}")
    print("  حساب مدیر بازرگانی -> یوزر: admin | پسورد: 123")
    print("  نمونه کارشناسان -> mansouri / tafaghodi / jamali / niroomand | پسورد: 123")
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
