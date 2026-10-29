"""Web wrapper for hde_ai_analysis.py.

Pick or upload an .xlsx file, run the script, watch the log, download the result.
Pure Python standard library (the script itself needs requests + openpyxl).
Login: one shared password from the APP_PASSWORD environment variable.
"""
import collections
import hashlib
import hmac
import importlib.util
import json
import mimetypes
import os
import re
import secrets
import shutil
import subprocess
import sys
import threading
import time
import traceback
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlsplit

DATA = Path(os.environ.get("DATA_ROOT", "/data"))
INBOX = DATA / "inbox"          # put .xlsx files here (WinSCP) or upload them on the page
WORK = DATA / "work"            # one folder per file: working copy + ticket cache
SCRIPT = Path(__file__).resolve().with_name("hde_ai_analysis.py")
STATIC = Path(__file__).resolve().with_name("static")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
BASE_PATH = os.environ.get("BASE_PATH", "").rstrip("/")
COOKIE_SECURE = os.environ.get("COOKIE_SECURE", "true").lower() == "true"
PORT = int(os.environ.get("PORT", "8000"))
COOKIE_NAME = "hde_sid"
SESSION_SECONDS = 30 * 86400
MAX_UPLOAD = 25 * 1024 * 1024


class ApiError(Exception):
    def __init__(self, status, msg):
        self.status, self.msg = status, msg


# ---------- sessions + login rate limit ----------
SESSIONS = {}           # sha256(token) -> expiry
FAILS = {}
SLOCK = threading.Lock()


def sha(s):
    return hashlib.sha256(s.encode()).hexdigest()


def rate_check(key, limit=8, window=300):
    now = time.time()
    FAILS[key] = [t for t in FAILS.get(key, []) if now - t < window]
    if len(FAILS[key]) >= limit:
        raise ApiError(429, "Too many attempts, try again in a few minutes")


# ---------- files ----------
def safe_name(raw):
    name = os.path.basename(str(raw).replace("\\", "/")).strip()
    name = re.sub(r"[^\w.\- ()]", "_", name)
    if not name.lower().endswith(".xlsx") or len(name) > 120 or name.startswith("."):
        raise ApiError(400, "Only .xlsx files are allowed (name up to 120 characters)")
    return name


def work_dir(name):
    return WORK / hashlib.sha1(name.encode()).hexdigest()[:16]


def result_info(name):
    meta = work_dir(name) / "meta.json"
    try:
        m = json.loads(meta.read_text())
        return m if m.get("ok") and (work_dir(name) / "input.xlsx").exists() else None
    except (OSError, ValueError):
        return None


def list_files():
    out = []
    for p in sorted(INBOX.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
        if p.name.startswith((".", "~$")):
            continue
        st = p.stat()
        res = result_info(p.name)
        out.append({"name": p.name, "size": st.st_size, "mtime": int(st.st_mtime),
                    "result_time": res["finished"] if res else None})
    return out


# ---------- job runner (one job at a time) ----------
JOB = {"state": "idle", "file": None, "started": None, "finished": None, "code": None}
LOG = collections.deque(maxlen=400)
LOCK = threading.Lock()
PROC = None
STOPPED = False


def _reader(proc, wdir):
    global PROC
    for line in proc.stdout:
        LOG.append(line.rstrip())
    code = proc.wait()
    with LOCK:
        JOB["code"] = code
        JOB["finished"] = int(time.time())
        if STOPPED:
            JOB["state"] = "stopped"
        elif code == 0:
            JOB["state"] = "done"
            (wdir / "meta.json").write_text(json.dumps({"ok": True, "finished": JOB["finished"]}))
        else:
            JOB["state"] = "error"
        PROC = None


def start_job(name, fresh):
    global PROC, STOPPED
    src = INBOX / name
    if not src.is_file():
        raise ApiError(404, "File not found")
    with LOCK:
        if JOB["state"] == "running":
            raise ApiError(409, "A job is already running")
        wdir = work_dir(name)
        wdir.mkdir(parents=True, exist_ok=True)
        (wdir / "meta.json").unlink(missing_ok=True)
        if fresh:
            (wdir / "input.cache.json").unlink(missing_ok=True)
        shutil.copyfile(src, wdir / "input.xlsx")        # script picks the only .xlsx in its folder
        # SOURCE_NAME: the working copy is always called input.xlsx, so pass the real file name
        # (e.g. queue_daily_20261025.xlsx): the script takes the conversation date from it
        env = dict(os.environ, DATA_DIR=str(wdir), SOURCE_NAME=name,
                   PYTHONUNBUFFERED="1", PYTHONIOENCODING="utf-8")
        env.pop("APP_PASSWORD", None)                       # the script does not need it
        LOG.clear()
        STOPPED = False
        PROC = subprocess.Popen([sys.executable, "-u", str(SCRIPT)], cwd=wdir, env=env,
                                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, encoding="utf-8", errors="replace")
        JOB.update(state="running", file=name, started=int(time.time()), finished=None, code=None)
        threading.Thread(target=_reader, args=(PROC, wdir), daemon=True).start()


def stop_job():
    global STOPPED
    with LOCK:
        if JOB["state"] == "running" and PROC:
            STOPPED = True
            PROC.terminate()


# ---------- API handlers ----------
def api_login(h, body, _):
    ip = h.client_ip()
    rate_check("login|" + ip)
    pw = body.get("password")
    ok = isinstance(pw, str) and hmac.compare_digest(sha(pw), sha(APP_PASSWORD))
    if not ok:
        FAILS.setdefault("login|" + ip, []).append(time.time())
        raise ApiError(400, "Wrong password")
    tok = secrets.token_urlsafe(32)
    with SLOCK:
        now = time.time()
        for k in [k for k, v in SESSIONS.items() if v < now]:
            del SESSIONS[k]
        SESSIONS[sha(tok)] = now + SESSION_SECONDS
    h.cookie(tok, SESSION_SECONDS)
    return {"ok": True}


def api_logout(h, body, _):
    tok = h.token()
    if tok:
        with SLOCK:
            SESSIONS.pop(sha(tok), None)
    h.cookie("", 0)
    return {"ok": True}


def api_files(h, body, _):
    return {"files": list_files(), "configured": bool(os.environ.get("HDE_EMAIL") and os.environ.get("HDE_API_KEY")),
            "domain": os.environ.get("HDE_DOMAIN", "kse"),
            "names_masked": importlib.util.find_spec("natasha") is not None}


def api_delete(h, body, groups):
    name = safe_name(unquote(groups[0]))
    p = INBOX / name
    if JOB["state"] == "running" and JOB["file"] == name:
        raise ApiError(409, "This file is being processed")
    if not p.is_file():
        raise ApiError(404, "File not found")
    p.unlink()
    shutil.rmtree(work_dir(name), ignore_errors=True)
    return {"ok": True}


def api_run(h, body, _):
    name = safe_name(body.get("name", ""))
    start_job(name, bool(body.get("fresh")))
    return {"ok": True}


def api_stop(h, body, _):
    stop_job()
    return {"ok": True}


def api_job(h, body, _):
    with LOCK:
        return dict(JOB, log="\n".join(list(LOG)[-200:]))


ROUTES = [
    ("POST", r"/api/login", api_login, False),
    ("POST", r"/api/logout", api_logout, False),
    ("GET", r"/api/me", lambda h, b, g: {"ok": True}, True),
    ("GET", r"/api/health", lambda h, b, g: {"ok": True}, False),
    ("GET", r"/api/files", api_files, True),
    ("DELETE", r"/api/files/(.+)", api_delete, True),
    ("POST", r"/api/run", api_run, True),
    ("POST", r"/api/stop", api_stop, True),
    ("GET", r"/api/job", api_job, True),
]
ROUTES = [(m, re.compile("^" + p + "$"), f, a) for m, p, f, a in ROUTES]

SEC_HEADERS = [
    ("X-Content-Type-Options", "nosniff"),
    ("Referrer-Policy", "same-origin"),
    ("X-Frame-Options", "DENY"),
    ("Content-Security-Policy",
     "default-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; "
     "frame-ancestors 'none'; base-uri 'none'; form-action 'self'"),
]


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "hde-analysis"

    def log_message(self, fmt, *args):
        sys.stdout.write("%s %s\n" % (self.address_string(), fmt % args))
        sys.stdout.flush()

    # --- helpers
    def client_ip(self):
        return self.headers.get("X-Forwarded-For", "").split(",")[0].strip() or self.client_address[0]

    def token(self):
        raw = self.headers.get("Cookie")
        if not raw:
            return None
        try:
            c = SimpleCookie(raw)
        except Exception:
            return None
        return c[COOKIE_NAME].value if COOKIE_NAME in c else None

    def authed(self):
        tok = self.token()
        with SLOCK:
            exp = SESSIONS.get(sha(tok)) if tok else None
        return bool(exp and exp > time.time())

    def cookie(self, value, max_age):
        parts = [f"{COOKIE_NAME}={value}", f"Max-Age={max_age}", f"Path={BASE_PATH or '/'}", "HttpOnly", "SameSite=Lax"]
        if COOKIE_SECURE:
            parts.append("Secure")
        self._cookies.append(("Set-Cookie", "; ".join(parts)))

    def _send(self, status, data, ctype, extra=()):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in list(extra) + SEC_HEADERS:
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _json(self, status, obj, extra=()):
        self._send(status, json.dumps(obj, ensure_ascii=False).encode(),
                   "application/json; charset=utf-8", list(extra) + [("Cache-Control", "no-store")])

    def _static(self, path):
        rel = "index.html" if path in ("", "/") else path.lstrip("/")
        full = (STATIC / rel).resolve()
        if STATIC.resolve() not in full.parents or not full.is_file():
            raise ApiError(404, "Not found")
        ctype = mimetypes.guess_type(str(full))[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype in ("application/json", "application/javascript"):
            ctype += "; charset=utf-8"
        self._send(200, full.read_bytes(), ctype, [("Cache-Control", "no-cache")])

    def _download(self, name):
        name = safe_name(name)
        if not result_info(name):
            raise ApiError(404, "No result yet. Run the file first")
        p = work_dir(name) / "input.xlsx"
        out_name = Path(name).stem + ".processed.xlsx"
        size = p.stat().st_size
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Length", str(size))
        self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + quote(out_name))
        self.send_header("Cache-Control", "no-store")
        for k, v in SEC_HEADERS:
            self.send_header(k, v)
        self.end_headers()
        with open(p, "rb") as f:
            shutil.copyfileobj(f, self.wfile)

    def _upload(self, name):
        name = safe_name(name)
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            self.close_connection = True
            raise ApiError(400, "Bad Content-Length")
        if length <= 0:
            raise ApiError(400, "Empty file")
        if length > MAX_UPLOAD:
            self.close_connection = True
            raise ApiError(413, f"File too large (max {MAX_UPLOAD // 1024 // 1024} MB)")
        if JOB["state"] == "running" and JOB["file"] == name:
            self.close_connection = True
            raise ApiError(409, "This file is being processed")
        tmp = INBOX / f".upload-{secrets.token_hex(6)}"
        try:
            left = length
            with open(tmp, "wb") as f:
                while left > 0:
                    chunk = self.rfile.read(min(65536, left))
                    if not chunk:
                        raise ApiError(400, "Upload interrupted")
                    f.write(chunk)
                    left -= len(chunk)
            with open(tmp, "rb") as f:
                if f.read(4) != b"PK\x03\x04":
                    raise ApiError(400, "This does not look like an .xlsx file")
            os.replace(tmp, INBOX / name)
        finally:
            tmp.unlink(missing_ok=True)
        return {"ok": True, "name": name}

    # --- request handling
    def _handle(self, method):
        self._cookies = []
        try:
            u = urlsplit(self.path)
            path = unquote(u.path)
            if not path.startswith("/api/"):
                if method != "GET":
                    raise ApiError(405, "Method not allowed")
                return self._static(path)

            if method != "GET":
                origin = self.headers.get("Origin")
                if origin and urlsplit(origin).netloc != self.headers.get("Host", ""):
                    raise ApiError(403, "Bad origin")

            m_up = re.match(r"^/api/files/(.+)$", path)
            if method == "PUT" and m_up:
                if not self.authed():
                    self.close_connection = True
                    raise ApiError(401, "Not signed in")
                return self._json(200, self._upload(m_up.group(1)))
            m_dl = re.match(r"^/api/result/(.+)$", path)
            if method == "GET" and m_dl:
                if not self.authed():
                    raise ApiError(401, "Not signed in")
                return self._download(m_dl.group(1))

            route = next(((fn, auth, mt.groups()) for m, rx, fn, auth in ROUTES
                          for mt in [rx.match(path)] if mt and m == method), None)
            try:
                n = int(self.headers.get("Content-Length") or 0)
            except ValueError:
                self.close_connection = True
                raise ApiError(400, "Bad Content-Length")
            if n > 100_000:
                self.close_connection = True
                raise ApiError(413, "Request too large")
            raw = self.rfile.read(n) if n else b""
            if not route:
                raise ApiError(404, "Not found")
            fn, auth, groups = route
            if auth and not self.authed():
                raise ApiError(401, "Not signed in")
            body = {}
            if raw.strip():
                try:
                    body = json.loads(raw)
                except ValueError:
                    raise ApiError(400, "Invalid JSON")
                if not isinstance(body, dict):
                    raise ApiError(400, "Invalid JSON")
            self._json(200, fn(self, body, groups), self._cookies)
        except ApiError as e:
            self._json(e.status, {"detail": e.msg})
        except (BrokenPipeError, ConnectionResetError):
            self.close_connection = True
        except Exception:
            traceback.print_exc()
            self._json(500, {"detail": "Server error"})

    def do_GET(self):
        self._handle("GET")

    def do_POST(self):
        self._handle("POST")

    def do_PUT(self):
        self._handle("PUT")

    def do_DELETE(self):
        self._handle("DELETE")


class Server(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


if __name__ == "__main__":
    if len(APP_PASSWORD) < 8:
        sys.exit("Set APP_PASSWORD in .env (at least 8 characters).")
    INBOX.mkdir(parents=True, exist_ok=True)
    WORK.mkdir(parents=True, exist_ok=True)
    print(f"hde-analysis listening on :{PORT}  data={DATA}", flush=True)
    Server(("0.0.0.0", PORT), Handler).serve_forever()
