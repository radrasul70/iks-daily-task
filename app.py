# -*- coding: utf-8 -*-
import http.server
import socketserver
import webbrowser
import socket
import json
import os
import ssl
import subprocess
import threading

PORT = int(os.environ.get("PORT", 8000))
HTTPS_PORT = int(os.environ.get("HTTPS_PORT", 8443))
IS_CLOUD = bool(os.environ.get("RENDER") or os.environ.get("RAILWAY_ENVIRONMENT") or os.environ.get("DYNO"))

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(DIRECTORY, "tasks_db.json")
CERT_FILE = os.path.join(DIRECTORY, "cert.pem")
KEY_FILE = os.path.join(DIRECTORY, "key.pem")
DB_LOCK = threading.Lock()

DEFAULT_USERS = [
    {
        "username": "admin",
        "password": "123",
        "name": "مدیر بازرگانی",
        "role": "مدیر بازرگانی"
    },
    {
        "username": "buyer1",
        "password": "123",
        "name": "کارشناس خرید ۱",
        "role": "کارشناس خرید"
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

def load_db():
    with DB_LOCK:
        if not os.path.exists(DB_FILE):
            initial = {
                "tasks": [],
                "nextNumber": 1,
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
                    return {
                        "tasks": data.get("tasks", []),
                        "nextNumber": data.get("nextNumber", 1),
                        "users": users
                    }
        except Exception:
            pass
        return {"tasks": [], "nextNumber": 1, "users": DEFAULT_USERS}

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
            for u in db["users"]:
                if u.get("username") == uname and u.get("password") == pwd:
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
            db["tasks"] = body.get("tasks", db.get("tasks", []))
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

    print("=" * 72)
    print("  برنامه دیلی تسک صوتی گروهی بازرگانی (نسخه 8 - ابری و محلی) فعال شد!")
    print(f"  1) آدرس در کامپیوتر: {pc_url}")
    if has_https:
        print(f"  2) آدرس امن در گوشی (روی همان Wi-Fi): {phone_https}")
    print("  حساب مدیر بازرگانی -> یوزر: admin | پسورد: 123")
    print("=" * 72)

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
