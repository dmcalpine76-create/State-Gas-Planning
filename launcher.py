"""
launcher.py - lets the State Gas control app start your planning tools on
this PC, so you press a button instead of typing into a Command Prompt.

It listens on this computer only (127.0.0.1:8765) and will only start a
fixed list of tools, only when asked by your control app's page, and only
with the pairing key held in launcher_key.txt. Each tool opens in its own
Command Prompt window, exactly as if you had typed the command yourself.

  py launcher.py            run it (install_launcher.bat starts it at login)
  py launcher.py --test     self-test, starts nothing
"""
import os
import sys
import json
import secrets
import datetime
import subprocess
import webbrowser
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from http.server import HTTPServer, BaseHTTPRequestHandler

HERE     = Path(__file__).resolve().parent
PORT     = 8765
APP_URL  = "https://dmcalpine76-create.github.io/stategas-planning/"
ALLOWED  = {"https://dmcalpine76-create.github.io"}
KEY_FILE = HERE / "launcher_key.txt"
LOG_FILE = HERE / "launcher.log"
VERSION  = "1"


def _python() -> str:
    """The console python.exe, even when this launcher runs windowless."""
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe" and (exe.parent / "python.exe").exists():
        return str(exe.parent / "python.exe")
    return str(exe)


# The only things the launcher will ever run. {py} = python, {days} = the
# lookback the app sent, checked to be a whole number from 1 to 120.
TOOLS = {
    "inbox":          {"title": "Inbox actions",
                       "cmd": '"{py}" inbox_actions.py run --days {days}'},
    "inbox_review":   {"title": "Inbox actions review",
                       "cmd": '"{py}" inbox_actions.py review'},
    "board_update":   {"title": "Weekly board update",
                       "cmd": '"{py}" board_update.py run --days {days}'},
    "board_briefing": {"title": "Monthly board briefing",
                       "cmd": '"{py}" board_briefing.py run --days {days} --docx'
                              ' && node generate_board_deck.js --data board_report_data.json'},
}


def load_key() -> str:
    if KEY_FILE.exists():
        k = KEY_FILE.read_text(encoding="utf-8").strip()
        if len(k) >= 32:
            return k
    k = secrets.token_hex(24)
    KEY_FILE.write_text(k, encoding="utf-8")
    return k


def log(msg: str):
    line = f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass
    if sys.stdout:
        try:
            print(line)
        except Exception:
            pass


def build_command(tool: str, days) -> str:
    spec = TOOLS.get(tool)
    if not spec:
        raise ValueError("unknown tool")
    if "{days}" in spec["cmd"]:
        try:
            d = int(days)
        except (TypeError, ValueError):
            raise ValueError("lookback must be a whole number of days")
        if not 1 <= d <= 120:
            raise ValueError("lookback must be between 1 and 120 days")
        return spec["cmd"].format(py=_python(), days=d)
    return spec["cmd"].format(py=_python())


def launch(tool: str, days) -> None:
    cmd = build_command(tool, days)
    title = TOOLS[tool]["title"]
    if os.name == "nt":
        # A new, visible Command Prompt that stays open when the tool ends,
        # so you can read what it did - same as running it by hand.
        subprocess.Popen(f'start "{title}" cmd /k "{cmd}"', cwd=HERE, shell=True)
    else:
        subprocess.Popen(cmd, cwd=HERE, shell=True)
    log(f"started {tool} ({cmd})")


class Handler(BaseHTTPRequestHandler):
    key = ""

    def log_message(self, *a):          # keep the console quiet
        pass

    def _origin_ok(self) -> bool:
        return self.headers.get("Origin", "") in ALLOWED

    def _cors(self):
        o = self.headers.get("Origin", "")
        if o in ALLOWED:
            self.send_header("Access-Control-Allow-Origin", o)
            self.send_header("Vary", "Origin")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type")
            # Chrome asks before a public site may call this PC; this says yes.
            self.send_header("Access-Control-Allow-Private-Network", "true")

    def _json(self, code: int, body: dict):
        data = json.dumps(body).encode("utf-8")
        self.send_response(code)
        self._cors()
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _host_ok(self) -> bool:
        # Blocks "DNS rebinding": a web address made to point at this PC.
        return self.headers.get("Host", "") in (f"127.0.0.1:{PORT}", f"localhost:{PORT}")

    def handle_one_request(self):
        try:
            return super().handle_one_request()
        except Exception as e:
            log(f"request error: {e}")

    def do_OPTIONS(self):
        if not self._host_ok():
            return self._json(403, {"ok": False})
        self.send_response(204 if self._origin_ok() else 403)
        self._cors()
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self):
        if not self._host_ok():
            return self._json(403, {"ok": False})
        path = urlparse(self.path).path
        if path == "/status":
            if not self._origin_ok():
                return self._json(403, {"ok": False})
            return self._json(200, {"ok": True, "version": VERSION,
                                    "tools": sorted(TOOLS)})
        if path == "/pair":
            # Opened by you from the app: sends you back to the app with the
            # key in the part of the address that never leaves your browser.
            # The key only ever travels to your own app page: no CORS header
            # is sent here, so another site calling this cannot read it.
            self.send_response(302)
            self.send_header("Location", f"{APP_URL}#pair={self.key}")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            return
        self._json(404, {"ok": False})

    def do_POST(self):
        if not self._host_ok():
            return self._json(403, {"ok": False})
        if urlparse(self.path).path != "/run" or not self._origin_ok():
            return self._json(403, {"ok": False, "error": "not allowed"})
        try:
            n = int(self.headers.get("Content-Length", "0"))
            body = json.loads(self.rfile.read(min(n, 4096)) or b"{}")
        except Exception:
            return self._json(400, {"ok": False, "error": "bad request"})
        if not secrets.compare_digest(str(body.get("key", "")), self.key):
            log("refused a run request with the wrong key")
            return self._json(401, {"ok": False, "error": "not paired - pair the app again"})
        try:
            launch(str(body.get("tool", "")), body.get("days"))
        except ValueError as e:
            return self._json(400, {"ok": False, "error": str(e)})
        except Exception as e:
            log(f"launch failed: {e}")
            return self._json(500, {"ok": False, "error": "could not start the tool"})
        return self._json(200, {"ok": True})


def self_test() -> int:
    ok = True
    for tool, days in (("inbox", 14), ("board_update", 7), ("board_briefing", 30), ("inbox_review", None)):
        c = build_command(tool, days)
        print(f"  {tool:15} {c}")
    for bad in (("inbox", "7; del *"), ("inbox", 0), ("inbox", 999), ("rm", 7)):
        try:
            build_command(*bad); print("  NOT REFUSED:", bad); ok = False
        except ValueError:
            pass
    print("  refuses bad input:", ok)
    return 0 if ok else 1


def main():
    if "--test" in sys.argv:
        sys.exit(self_test())
    Handler.key = load_key()
    try:
        server = HTTPServer(("127.0.0.1", PORT), Handler)
    except OSError:
        log(f"launcher already running on port {PORT} - nothing to do")
        if "--pair" in sys.argv:
            webbrowser.open(f"http://127.0.0.1:{PORT}/pair")
        return
    log(f"launcher v{VERSION} listening on 127.0.0.1:{PORT}")
    if "--pair" in sys.argv:
        webbrowser.open(f"http://127.0.0.1:{PORT}/pair")
    server.serve_forever()


if __name__ == "__main__":
    main()
