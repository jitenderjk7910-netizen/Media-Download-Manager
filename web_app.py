"""Web dashboard for Media Download Manager.

Serves a single-page dashboard (loaded from templates/index.html) over the
built-in ThreadingHTTPServer. All download orchestration lives in
universal_link_downloader.UniversalDownloader; this module wires the HTTP API,
job state, browser fallback, search providers, and job-config persistence.
"""

import csv
import hashlib
import hmac
import html
import json
import logging
import mimetypes
import os
import re
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.parse
import webbrowser
import zipfile
from email.parser import BytesParser
from email.policy import default as email_policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
try:
    from tkinter import Tk, filedialog
    TKINTER_AVAILABLE = True
except (ImportError, Exception):
    TKINTER_AVAILABLE = False

import requests

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from universal_link_downloader import (
    ARCHIVE_EXTENSIONS,
    DownloadResult,
    IMAGE_EXTENSIONS,
    LinkTask,
    LOG_NAME,
    UniversalDownloader,
    VIDEO_EXTENSIONS,
    is_internal_metadata_file,
    open_in_file_manager,
    provider_name,
    rows_from_file,
    rows_from_pasted_text,
    sanitize_name,
    tasks_from_rows,
    unique_path,
)

# Centralized configuration (defaults + optional config.json overrides).
from app_config import (
    APP_VERSION,
    ALLOWED_ORIGINS,
    AUTH_PROFILE,
    AUTH_PASSWORD,
    AUTH_USERNAME,
    BROWSER_CLICK_TIMEOUT_MS,
    BROWSER_CONTROL_TIMEOUT_MS,
    BROWSER_DOWNLOAD_TIMEOUT_MS,
    BROWSER_IMAGE_CHECK_TIMEOUT_MS,
    BROWSER_PAGE_TIMEOUT_MS,
    BROWSER_SCREENSHOT_TIMEOUT_MS,
    DEFAULT_NOTIFICATIONS,
    DEFAULT_OUTPUT,
    DEFAULT_PROXY,
    DEFAULT_SPEED_LIMIT_KBPS,
    DEFAULT_TIMEOUT,
    DEFAULT_WORKERS,
    EFFECTIVE_HOST,
    EFFECTIVE_NOTIFICATIONS,
    EFFECTIVE_PORT,
    EFFECTIVE_PROXY,
    EFFECTIVE_SPEED_LIMIT_KBPS,
    EFFECTIVE_WORKERS,
    MAX_BROWSER_IMAGES,
    MIN_BROWSER_IMAGE_AREA,
    MAX_LOG_LINES,
    MAX_PREVIEW_SAMPLE,
    MAX_REPORT_ROWS,
    MAX_THUMB_SIZE,
    PORT_FALLBACK_LIMIT,
    PRODUCTION,
    PROVIDER_LIMITS,
    RATE_LIMIT_REQUESTS,
    RATE_LIMIT_WINDOW,
    SAVED_CONFIGS_DIR,
    SESSION_COOKIE_NAME,
    SESSION_MAX_AGE,
    SESSION_SECRET,
    TEMPLATES_DIR,
    THUMB_EXTENSIONS,
    TUNNEL_URL,
    UPLOAD_DIR,
)

# HOST/PORT: env vars take priority (already resolved in app_config).
HOST = EFFECTIVE_HOST
PORT = EFFECTIVE_PORT

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------
if PRODUCTION:
    logging.basicConfig(
        filename=str(APP_DIR / "app.log"),
        level=logging.INFO,
        format='{"time":"%(asctime)s","level":"%(levelname)s","msg":"%(message)s"}',
        datefmt="%Y-%m-%dT%H:%M:%S",
    )
else:
    logging.basicConfig(level=logging.WARNING)

logger = logging.getLogger("mdm")

# Ensure runtime directories exist.
DEFAULT_OUTPUT.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
SAVED_CONFIGS_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------
AUTH_ENABLED = bool(AUTH_USERNAME and AUTH_PASSWORD)

def _make_session_token(username: str) -> str:
    """Create an HMAC-signed session token."""
    payload = f"{username}:{int(time.time()) // SESSION_MAX_AGE}"
    sig = hmac.new(SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    return f"{payload}:{sig}"

def _validate_session_token(token: str) -> bool:
    """Verify an HMAC-signed session token is valid and not expired."""
    try:
        parts = token.rsplit(":", 1)
        if len(parts) != 2:
            return False
        payload, sig = parts
        expected = hmac.new(SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            return False
        # Check time window
        _username, ts_bucket = payload.rsplit(":", 1)
        current_bucket = int(time.time()) // SESSION_MAX_AGE
        return int(ts_bucket) >= current_bucket - 1  # allow 1 bucket overlap
    except Exception:
        return False

def _get_session_cookie(handler) -> str:
    """Extract the session cookie value from request headers."""
    cookie_header = handler.headers.get("Cookie", "")
    for part in cookie_header.split(";"):
        part = part.strip()
        if part.startswith(f"{SESSION_COOKIE_NAME}="):
            return part[len(SESSION_COOKIE_NAME) + 1:]
    return ""

def _is_authenticated(handler) -> bool:
    """Return True if auth is disabled or the session token is valid."""
    if not AUTH_ENABLED:
        return True
    token = _get_session_cookie(handler)
    return bool(token) and _validate_session_token(token)

# ---------------------------------------------------------------------------
# Rate limiting (token bucket per IP)
# ---------------------------------------------------------------------------
_rate_buckets: dict = {}
_rate_lock = threading.Lock()

def _check_rate_limit(ip: str) -> bool:
    """Return True if request is allowed, False if rate limited."""
    now = time.time()
    with _rate_lock:
        bucket = _rate_buckets.get(ip)
        if bucket is None or now - bucket["window_start"] >= RATE_LIMIT_WINDOW:
            _rate_buckets[ip] = {"window_start": now, "count": 1}
            return True
        bucket["count"] += 1
        return bucket["count"] <= RATE_LIMIT_REQUESTS



# ---------------------------------------------------------------------------
# Job state (thread-safe)
# ---------------------------------------------------------------------------
class JobState:
    def __init__(self):
        self.lock = threading.Lock()
        self.reset()

    def reset(self):
        self.start_time = 0
        self.running = False
        self.done = 0
        self.total = 0
        self.ok = 0
        self.failed = 0
        self.skipped = 0
        self.files = 0
        self.output = str(DEFAULT_OUTPUT)
        self.report = ""
        self.status = "Ready"
        self.providers = {}
        self.failure_groups = {}
        self.retry_queue = {}
        self.audit = {}
        self.bytes_downloaded = 0
        self.active_links = {}
        self.last_activity = ""
        self.logs = []
        self.rows = []
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.notify_enabled = EFFECTIVE_NOTIFICATIONS
        self.downloader = None

    def snapshot(self):
        with self.lock:
            elapsed = max(0, time.time() - self.start_time) if self.start_time else 0
            remaining = max(0, self.total - self.done)
            links_per_min = (self.done / elapsed * 60) if elapsed > 0 else 0
            files_per_min = (self.files / elapsed * 60) if elapsed > 0 else 0
            mb_per_sec = (self.bytes_downloaded / 1024 / 1024 / elapsed) if elapsed > 0 else 0
            eta_seconds = int(remaining / (self.done / elapsed)) if self.done and elapsed > 0 else 0
            return {
                "start_time": self.start_time,
                "elapsedSeconds": int(elapsed),
                "etaSeconds": eta_seconds,
                "running": self.running,
                "done": self.done,
                "total": self.total,
                "ok": self.ok,
                "failed": self.failed,
                "skipped": self.skipped,
                "files": self.files,
                "bytesDownloaded": self.bytes_downloaded,
                "mbPerSec": round(mb_per_sec, 2),
                "linksPerMin": round(links_per_min, 1),
                "filesPerMin": round(files_per_min, 1),
                "output": self.output,
                "report": self.report,
                "status": self.status,
                "paused": self.pause_event.is_set(),
                "providers": self.providers,
                "failureGroups": self.failure_groups,
                "retryQueue": self.retry_queue,
                "audit": self.audit,
                "activeLinks": list(self.active_links.values())[:12],
                "lastActivity": self.last_activity,
                "logs": self.logs[-MAX_LOG_LINES:],
                "rows": self.rows[-MAX_REPORT_ROWS:],
                "notifyEnabled": self.notify_enabled,
            }

    def update(self, **kwargs):
        with self.lock:
            for key, value in kwargs.items():
                setattr(self, key, value)

    def mark_started(self):
        with self.lock:
            self.start_time = time.time()

    def log(self, text):
        with self.lock:
            self.logs.append(str(text))
            self.logs = self.logs[-MAX_LOG_LINES:]

    def add_row(self, item):
        with self.lock:
            self.rows.append(item)
            self.rows = self.rows[-MAX_REPORT_ROWS:]

    def add_bytes(self, count, url="", task=None):
        with self.lock:
            count = int(count or 0)
            self.bytes_downloaded += count
            if task is not None:
                key = f"{task.row}:{task.url}"
                if key in self.active_links:
                    self.active_links[key]["bytes"] = self.active_links[key].get("bytes", 0) + count
            if url:
                self.last_activity = f"Receiving: {url[:120]}"

    def task_event(self, event, task, provider, attempt, result=None):
        key = f"{task.row}:{task.url}"
        with self.lock:
            if event == "start":
                self.active_links[key] = {
                    "row": task.row,
                    "folder": task.folder,
                    "provider": provider,
                    "attempt": attempt,
                    "url": task.url[:120],
                    "bytes": 0,
                }
                self.last_activity = f"Row {task.row} {provider} attempt {attempt}"
            else:
                self.active_links.pop(key, None)


STATE = JobState()


# ---------------------------------------------------------------------------
# Helpers: HTML template, desktop notifications, file pickers
# ---------------------------------------------------------------------------
def _load_login_html(error: str = "") -> str:
    """Generate a clean built-in login page (no Supabase dependency)."""
    err_html = f'<p class="error">{html.escape(error)}</p>' if error else ""
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Media Download Manager — Login</title>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    min-height: 100vh; display: flex; align-items: center; justify-content: center;
    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
    font-family: 'Segoe UI', system-ui, sans-serif; color: #fff;
  }}
  .card {{
    background: rgba(255,255,255,0.07); backdrop-filter: blur(16px);
    border: 1px solid rgba(255,255,255,0.15); border-radius: 20px;
    padding: 48px 40px; width: 100%; max-width: 400px;
    box-shadow: 0 25px 50px rgba(0,0,0,0.5);
  }}
  .logo {{ font-size: 2rem; font-weight: 700; margin-bottom: 8px; }}
  .logo span {{ color: #a78bfa; }}
  .subtitle {{ color: rgba(255,255,255,0.55); font-size: 0.9rem; margin-bottom: 32px; }}
  label {{ display: block; font-size: 0.8rem; font-weight: 600; letter-spacing: .05em;
    color: rgba(255,255,255,0.7); margin-bottom: 6px; }}
  input {{
    width: 100%; padding: 12px 16px; background: rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.2); border-radius: 10px;
    color: #fff; font-size: 1rem; outline: none; margin-bottom: 20px;
    transition: border-color .2s;
  }}
  input:focus {{ border-color: #a78bfa; }}
  button {{
    width: 100%; padding: 13px; background: linear-gradient(135deg, #7c3aed, #a78bfa);
    border: none; border-radius: 10px; color: #fff; font-size: 1rem;
    font-weight: 600; cursor: pointer; transition: opacity .2s;
  }}
  button:hover {{ opacity: 0.88; }}
  .error {{ color: #f87171; font-size: 0.88rem; margin-bottom: 16px;
    background: rgba(248,113,113,0.1); border-radius: 8px; padding: 10px 14px; }}
</style>
</head>
<body>
<div class="card">
  <div class="logo">Media <span>DL</span></div>
  <div class="subtitle">Sign in to access the Download Manager</div>
  {err_html}
  <form method="POST" action="/auth/login">
    <label for="username">Username</label>
    <input id="username" name="username" type="text" autocomplete="username" required autofocus>
    <label for="password">Password</label>
    <input id="password" name="password" type="password" autocomplete="current-password" required>
    <button type="submit">Sign In</button>
  </form>
</div>
</body>
</html>"""

def pick_output() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askdirectory(title="Select Output Folder")
        root.destroy()
        return path or ""
    except Exception:
        return ""

def _load_html() -> str:
    """Load the frontend from templates/index.html (kept out of this module)."""
    template_path = TEMPLATES_DIR / "index.html"
    if template_path.exists():
        return template_path.read_text(encoding="utf-8")
    return (
        '<!DOCTYPE html><html><body style="font-family:sans-serif;padding:40px">'
        "<h1>Frontend template missing</h1>"
        "<p>Put <code>index.html</code> in the <code>templates/</code> folder.</p>"
        "</body></html>"
    )


def _send_desktop_notification(title: str, body: str) -> None:
    """Best-effort OS desktop notification (Windows/macOS/Linux)."""
    try:
        if sys.platform.startswith("win"):
            ps = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "$n = New-Object System.Windows.Forms.NotifyIcon; "
                "$n.Icon = [System.Drawing.SystemIcons]::Information; "
                "$n.Visible = $true; "
                f"$n.ShowBalloonTip(5000, '{title}', '{body}', 'Info'); "
                "Start-Sleep -Seconds 6; $n.Dispose()"
            )
            subprocess.Popen(["powershell", "-NoProfile", "-Command", ps],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        elif sys.platform == "darwin":
            safe_body = body.replace('"', '\\"')
            safe_title = title.replace('"', '\\"')
            subprocess.Popen(["osascript", "-e",
                              f'display notification "{safe_body}" with title "{safe_title}"'],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            subprocess.Popen(["notify-send", title, body],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass  # Best-effort; never fail the main flow.


def pick_file():
    if not TKINTER_AVAILABLE:
        raise RuntimeError("File picker is not available in cloud/server mode. Please type the path manually.")
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="Select Excel or CSV file",
        filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")],
    )
    root.destroy()
    return path


def pick_output():
    if not TKINTER_AVAILABLE:
        raise RuntimeError("Folder picker is not available in cloud/server mode. Please type the path manually.")
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askdirectory(title="Select output folder")
    root.destroy()
    return path


def open_path(path):
    value = Path(path or DEFAULT_OUTPUT)
    value.mkdir(parents=True, exist_ok=True)
    open_in_file_manager(value.resolve())


def create_excel_template(path):
    from openpyxl import Workbook
    folder = Path(path or DEFAULT_OUTPUT)
    folder.mkdir(parents=True, exist_ok=True)
    template = folder / "link_download_template.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Links"
    ws.append(["Folder / Barcode", "Link 1", "Link 2", "Link 3"])
    ws.append(["ABC001", "https://example.com/image.jpg", "https://www.dropbox.com/scl/fo/...", ""])
    ws.append(["ABC002", "https://drive.google.com/open?id=...", "", ""])
    for col, width in {"A": 24, "B": 48, "C": 48, "D": 48}.items():
        ws.column_dimensions[col].width = width
    wb.save(template)
    open_in_file_manager(template.resolve())


# ---------------------------------------------------------------------------
# Thumbnails, classification, audit
# ---------------------------------------------------------------------------
def first_thumbnail(folder):
    folder = Path(folder)
    if not folder.exists():
        return ""
    for path in folder.rglob("*"):
        if (
            path.is_file()
            and ".partials" not in path.parts
            and not is_internal_metadata_file(path)
            and path.suffix.lower() in IMAGE_EXTENSIONS
        ):
            return str(path.resolve())
    return ""


def classify_failure(status, message):
    text = f"{status} {message}".lower()
    if status == "ok":
        return ""
    if status.startswith("skipped"):
        return "skipped"
    if "login" in text or "permission" in text or "access" in text or "sign in" in text:
        return "login/access"
    if "404" in text or "not found" in text:
        return "not found"
    if "timeout" in text or "timed out" in text:
        return "timeout"
    if "network" in text or "connection" in text or "reset" in text:
        return "network"
    if "429" in text or "rate" in text:
        return "rate limited"
    if any(code in text for code in ("500", "502", "503", "504")):
        return "server"
    if "corrupt" in text or "empty" in text or "signature" in text:
        return "invalid file"
    if "no image" in text or "no downloadable" in text:
        return "no media"
    if "html" in text or "page" in text:
        return "page/not direct"
    return "other"


def format_seconds(seconds):
    seconds = max(0, int(seconds or 0))
    minutes, secs = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    if hours:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


def output_audit(output, expected_tasks):
    output = Path(output)
    expected_folders = {task.folder for task in expected_tasks}
    empty = []
    existing = 0
    for folder in expected_folders:
        path = output / sanitize_name(folder)
        files = [
            item
            for item in path.rglob("*")
            if item.is_file() and ".partials" not in item.parts and not is_internal_metadata_file(item)
        ] if path.exists() else []
        if files:
            existing += 1
        else:
            empty.append(folder)
    summary = f"Audit: {existing}/{len(expected_folders)} folders contain files"
    if empty:
        summary += f"; empty/missing {len(empty)}"
    return {"summary": summary, "empty": empty[:50]}


def cleanup_output_metadata(output):
    output = Path(output)
    if not output.exists():
        return 0
    removed = 0
    for path in output.rglob("*"):
        if path.is_file() and is_internal_metadata_file(path):
            path.unlink(missing_ok=True)
            removed += 1
    return removed


def is_rename_candidate(path):
    if not path.is_file():
        return False
    if ".partials" in path.parts or is_internal_metadata_file(path):
        return False
    if path.name == LOG_NAME:
        return False
    suffix = path.suffix.lower()
    return suffix in IMAGE_EXTENSIONS or suffix in VIDEO_EXTENSIONS or suffix in ARCHIVE_EXTENSIONS


def rename_downloaded_files(output):
    output = Path(output or DEFAULT_OUTPUT)
    if not output.exists() or not output.is_dir():
        raise RuntimeError("Output folder not found.")

    removed_metadata = cleanup_output_metadata(output)
    renamed = 0
    folders = 0
    skipped = 0
    for product_dir in sorted([p for p in output.iterdir() if p.is_dir()], key=lambda p: p.name.lower()):
        if product_dir.name.startswith("."):
            continue
        files = sorted(
            [p for p in product_dir.rglob("*") if is_rename_candidate(p)],
            key=lambda p: str(p.relative_to(product_dir)).lower(),
        )
        if not files:
            continue
        folders += 1
        base = sanitize_name(product_dir.name)
        temp_pairs = []
        for index, src in enumerate(files, start=1):
            temp = src.with_name(f".__rename_tmp_{time.time_ns()}_{index}{src.suffix}")
            src.rename(temp)
            temp_pairs.append((temp, src.parent / f"{base}_{index}{src.suffix.lower()}"))
        for temp, final_path in temp_pairs:
            if final_path.exists():
                final_path = unique_path(final_path.parent, final_path.name)
            if temp.name == final_path.name:
                skipped += 1
                temp.rename(final_path)
                continue
            temp.rename(final_path)
            renamed += 1
    return {"renamed": renamed, "folders": folders, "skipped": skipped, "removedMetadata": removed_metadata}


# ---------------------------------------------------------------------------
# Browser fallback (Playwright)
# ---------------------------------------------------------------------------
def save_browser_download(download, target):
    target = Path(target)
    target.mkdir(parents=True, exist_ok=True)
    suggested = sanitize_name(download.suggested_filename or "browser_download")
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir) / suggested
        download.save_as(str(tmp_path))
        helper = UniversalDownloader(target)
        if zipfile.is_zipfile(tmp_path):
            count = helper.extract_archive(tmp_path, target)
            if count:
                return DownloadResult("ok", count, f"Browser extracted {count} file(s)")
            return DownloadResult("no_files", 0, "Browser downloaded zip, but no usable files found.")
        dest = unique_path(target, suggested)
        tmp_path.replace(dest)
        return DownloadResult("ok", 1, f"Browser saved {dest.name}")


def save_largest_visible_image(page, target):
    target = Path(target)
    images = page.locator("img")
    best = None
    best_area = 0
    for index in range(min(images.count(), MAX_BROWSER_IMAGES)):
        locator = images.nth(index)
        try:
            box = locator.bounding_box(timeout=BROWSER_IMAGE_CHECK_TIMEOUT_MS)
        except Exception:
            box = None
        if not box:
            continue
        area = float(box.get("width", 0)) * float(box.get("height", 0))
        if area > best_area:
            best = locator
            best_area = area
    if best is None or best_area < MIN_BROWSER_IMAGE_AREA:
        return DownloadResult("failed", 0, "Browser page opened, but no usable image/download control was found.")
    target.mkdir(parents=True, exist_ok=True)
    dest = unique_path(target, "browser_image.png")
    best.screenshot(path=str(dest), timeout=BROWSER_SCREENSHOT_TIMEOUT_MS)
    return DownloadResult("ok", 1, f"Browser captured visible image: {dest.name}")


def browser_fallback_download(tasks, output, progress):
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        STATE.log(f"Browser fallback unavailable: {exc}. Run install_requirements.bat.")
        return []

    results = []
    profile = AUTH_PROFILE
    profile.mkdir(parents=True, exist_ok=True)
    selectors = [
        'button[title*="Download" i]',
        'button[aria-label*="Download" i]',
        'div[role="button"][aria-label*="Download" i]',
        '[title*="Download" i]',
        '[aria-label*="Download" i]',
        '[data-automationid*="download" i]',
        '[data-testid*="download" i]',
        'button[name*="download" i]',
        'a[href*="download" i]',
        'i[data-icon-name*="Download" i]',
        'span:has-text("Download")',
        'button:has-text("Download")',
    ]
    with sync_playwright() as playwright:
        context = None
        last_error = None
        for channel in ("chrome", "msedge"):
            try:
                context = playwright.chromium.launch_persistent_context(
                    str(profile.resolve()),
                    channel=channel,
                    headless=True,
                    accept_downloads=True,
                    viewport={"width": 1366, "height": 850},
                )
                break
            except Exception as exc:
                last_error = exc
        if context is None:
            STATE.log(f"Browser fallback could not open Chrome/Edge: {last_error}")
            return []
        page = context.pages[0] if context.pages else context.new_page()
        try:
            for task in tasks:
                target = Path(output) / sanitize_name(task.folder)
                result = DownloadResult("failed", 0, "Browser fallback did not find a download button.")
                try:
                    page.goto(task.url, wait_until="domcontentloaded", timeout=BROWSER_PAGE_TIMEOUT_MS)
                    url_lower = page.url.lower()
                    if ("login.live.com" in url_lower or "login.microsoftonline.com" in url_lower) and "onedrive.live.com" not in url_lower:
                        result = DownloadResult("login_required", 0, "Browser opened login page. Login in the browser and retry.")
                    else:
                        clicked = False
                        possible_locators = [
                            page.get_by_role("button", name=re.compile("download", re.IGNORECASE)),
                            page.get_by_label(re.compile("download", re.IGNORECASE)),
                            page.get_by_title(re.compile("download", re.IGNORECASE)),
                        ]
                        for selector in selectors:
                            possible_locators.append(page.locator(selector))

                        for locator_collection in possible_locators:
                            locator = locator_collection.first
                            try:
                                locator.wait_for(state="visible", timeout=BROWSER_CONTROL_TIMEOUT_MS)
                                with page.expect_download(timeout=BROWSER_DOWNLOAD_TIMEOUT_MS) as download_info:
                                    locator.click(timeout=BROWSER_CLICK_TIMEOUT_MS)
                                result = save_browser_download(download_info.value, target)
                                clicked = True
                                break
                            except PlaywrightTimeoutError:
                                continue
                            except Exception:
                                continue
                        if not clicked:
                            result = save_largest_visible_image(page, target)
                except Exception as exc:
                    result = DownloadResult("failed", 0, f"Browser fallback failed: {exc}")
                results.append((task, result))
                progress(task, result)
        finally:
            context.close()
    return results


# ---------------------------------------------------------------------------
# Search providers
# ---------------------------------------------------------------------------
def search_terms_from_payload(payload):
    path = (payload.get("searchInputPath") or payload.get("inputPath") or "").strip().strip('"')
    terms = []
    if path:
        for row in rows_from_file(path):
            value = first_search_term_from_row(row)
            if value:
                terms.append(value)

    text = payload.get("searchTerms") or ""
    for line in text.splitlines():
        value = clean_style_code(line)
        if value and not is_search_header(value):
            terms.append(value)
    return list(dict.fromkeys(terms))


def clean_style_code(value):
    value = str(value or "").strip()
    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]
    return value


def first_search_term_from_row(row):
    for cell in row:
        value = clean_style_code(cell)
        if not value:
            continue
        if is_search_header(value):
            return ""
        return value
    return ""


def is_search_header(value):
    lowered = str(value or "").strip().lower().replace("_", " ")
    return lowered in {"style code", "style", "code", "sku", "barcode", "search term", "search"}


def image_links_from_text(text):
    found = []
    seen = set()
    for match in re.finditer(r"https?://[^\"'<>\s]+?\.(?:jpg|jpeg|png|webp)(?:\?[^\"'<>\s]*)?", text, flags=re.I):
        url = match.group(0).replace("\\/", "/")
        if url not in seen:
            seen.add(url)
            found.append(url)
    return found


def clean_discovered_image_url(url):
    url = str(url or "").strip().strip('"').strip("'")
    if not url:
        return ""
    url = url.replace("\\/", "/")
    url = url.replace("\\u003d", "=").replace("\\u0026", "&")
    url = html.unescape(url)
    if url.startswith("/imgres?"):
        parsed = urllib.parse.urlsplit(url)
        params = dict(urllib.parse.parse_qsl(parsed.query))
        url = params.get("imgurl") or ""
    if url.startswith("http"):
        return url
    return ""


def image_links_from_google_html(text):
    decoded = text.encode("utf-8", "ignore").decode("unicode_escape", "ignore")
    candidates = []
    patterns = [
        r'"ou"\s*:\s*"([^"]+)"',
        r'imgurl=([^&"\'>\s]+)',
        r'(https?://[^"\'<>\s\\]+?\.(?:jpg|jpeg|png|webp)(?:\?[^"\'<>\s\\]*)?)',
    ]
    for source in (text, decoded):
        for pattern in patterns:
            for match in re.finditer(pattern, source, flags=re.I):
                url = clean_discovered_image_url(urllib.parse.unquote(match.group(1)))
                if url:
                    candidates.append(url)
    filtered = []
    blocked_hosts = ("google.com", "gstatic.com/images/branding", "schema.org")
    for url in candidates:
        lowered = url.lower()
        if any(host in lowered for host in blocked_hosts):
            continue
        if url not in filtered:
            filtered.append(url)
    return filtered


def prefer_product_images(urls, quality_mode="best"):
    product_urls = [url for url in urls if "/products/" in url]
    candidates = product_urls or urls
    if quality_mode == "all":
        return list(dict.fromkeys(candidates))
    sized = []
    for url in candidates:
        size_match = re.search(r"_size(\d+)x(\d+)", url)
        if size_match:
            area = int(size_match.group(1)) * int(size_match.group(2))
            sized.append((area, url))
    if any(area >= 150000 for area, _ in sized):
        candidates = [url for area, url in sized if area >= 150000]
    chosen = {}
    for url in candidates:
        key = re.sub(r"_size\d+x\d+(?=\.)", "", url)
        size_match = re.search(r"_size(\d+)x(\d+)", url)
        score = 0
        if size_match:
            score = int(size_match.group(1)) * int(size_match.group(2))
        if key not in chosen or score > chosen[key][0]:
            chosen[key] = (score, url)
    return [item[1] for item in chosen.values()]


def search_koton_with_requests(term, quality_mode="best"):
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36"
        )
    })
    search_url = f"https://www.koton.com/list/?search_text={urllib.parse.quote(term)}"
    response = session.get(search_url, timeout=30)
    response.raise_for_status()
    product_urls = []
    for match in re.finditer(r'href=["\']([^"\']+)["\']', response.text, flags=re.I):
        href = urllib.parse.urljoin(search_url, match.group(1))
        if "koton.com" in href and href not in product_urls and "/p/" in href:
            product_urls.append(href)
    if not product_urls:
        return []

    images = []
    for product_url in product_urls[:1]:
        product = session.get(product_url, timeout=30)
        product.raise_for_status()
        images.extend(image_links_from_text(product.text))
    return prefer_product_images(images, quality_mode=quality_mode)


def search_google_images_with_requests(term, quality_mode="best"):
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
    })
    query = f'"{term}"'
    search_url = "https://www.google.com/search?" + urllib.parse.urlencode({
        "tbm": "isch",
        "safe": "off",
        "q": query,
    })
    response = session.get(search_url, timeout=30)
    response.raise_for_status()
    images = image_links_from_google_html(response.text)
    if quality_mode != "all":
        return images[:1]
    return images[:8]


def search_google_images_with_selenium(term, quality_mode="best"):
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.common.by import By
        from webdriver_manager.chrome import ChromeDriverManager
    except Exception as exc:
        raise RuntimeError("Selenium dependencies missing. Run install_requirements.bat.") from exc

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1366,900")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    try:
        search_url = "https://www.google.com/search?" + urllib.parse.urlencode({
            "tbm": "isch",
            "safe": "off",
            "q": f'"{term}"',
        })
        driver.get(search_url)
        driver.implicitly_wait(4)
        urls = []
        for image in driver.find_elements(By.CSS_SELECTOR, "img"):
            for attr in ("src", "data-src", "data-iurl"):
                value = clean_discovered_image_url(image.get_attribute(attr))
                if value and value.startswith("http"):
                    urls.append(value)
        urls = list(dict.fromkeys(urls))
        if quality_mode != "all":
            return urls[:1]
        return urls[:8]
    finally:
        driver.quit()


def search_google_tasks(terms, log=None, quality_mode="best"):
    tasks = []
    for row, term in enumerate(terms, start=1):
        if log:
            log(f"Google Images search: {term}")
        try:
            images = search_google_images_with_requests(term, quality_mode=quality_mode)
            if not images:
                images = search_google_images_with_selenium(term, quality_mode=quality_mode)
        except Exception as exc:
            if log:
                log(f"Google search failed for {term}: {exc}")
            images = []
        for image_url in images:
            tasks.append(LinkTask(row=row, folder=term, url=image_url))
    return tasks


def search_koton_with_selenium(term, quality_mode="best"):
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.common.by import By
        from webdriver_manager.chrome import ChromeDriverManager
    except Exception as exc:
        raise RuntimeError("Selenium dependencies missing. Run install_requirements.bat.") from exc

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1366,900")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    try:
        search_url = f"https://www.koton.com/list/?search_text={urllib.parse.quote(term)}"
        driver.get(search_url)
        driver.implicitly_wait(4)
        product_links = driver.find_elements(By.CSS_SELECTOR, ".product-item a, a[href*='/p/']")
        if product_links:
            href = product_links[0].get_attribute("href")
            if href:
                driver.get(href)
        driver.implicitly_wait(4)
        image_elements = driver.find_elements(By.CSS_SELECTOR, "img")
        urls = []
        for image in image_elements:
            for attr in ("src", "data-src", "data-original"):
                value = image.get_attribute(attr)
                if value and value.startswith("http") and re.search(r"\.(jpg|jpeg|png|webp)(?:\?|$)", value, flags=re.I):
                    urls.append(value)
        return prefer_product_images(list(dict.fromkeys(urls)), quality_mode=quality_mode)
    finally:
        driver.quit()


def search_tasks_from_terms(terms, log=None, quality_mode="best"):
    tasks = []
    for row, term in enumerate(terms, start=1):
        if log:
            log(f"Searching Koton: {term}")
        try:
            images = search_koton_with_requests(term, quality_mode=quality_mode)
            if not images:
                images = search_koton_with_selenium(term, quality_mode=quality_mode)
        except Exception as exc:
            if log:
                log(f"Search failed for {term}: {exc}")
            images = []
        for image_url in images:
            tasks.append(LinkTask(row=row, folder=term, url=image_url))
    return tasks


def search_custom_tasks(terms, template, log=None, quality_mode="best"):
    if "{term}" not in template:
        raise RuntimeError("Custom search template must contain {term}.")
    tasks = []
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    for row, term in enumerate(terms, start=1):
        page_url = template.replace("{term}", urllib.parse.quote(term))
        if log:
            log(f"Scanning custom page: {page_url}")
        response = session.get(page_url, timeout=30)
        response.raise_for_status()
        for image_url in prefer_product_images(image_links_from_text(response.text), quality_mode=quality_mode):
            tasks.append(LinkTask(row=row, folder=term, url=image_url))
    return tasks


# ---------------------------------------------------------------------------
# Task building / resume / preview
# ---------------------------------------------------------------------------
def retry_failed_tasks(output):
    report = Path(output or DEFAULT_OUTPUT) / LOG_NAME
    if not report.exists():
        raise RuntimeError(f"No {LOG_NAME} found in selected output folder.")
    tasks = []
    rows = rows_from_file(report)
    if not rows:
        raise RuntimeError("Report is empty.")
    headers = [str(cell).lower().strip() for cell in rows[0]]
    for row_data in rows[1:]:
        row_data = list(row_data) + [""] * (len(headers) - len(row_data))
        row = dict(zip(headers, row_data))
        status = str(row.get("status") or "").lower()
        url = str(row.get("url") or "").strip()
        if status and status != "ok" and not status.startswith("skipped") and url.startswith(("http://", "https://")):
            row_num = row.get("row")
            try:
                row_num = int(row_num)
            except (TypeError, ValueError):
                row_num = len(tasks) + 1
            tasks.append(LinkTask(
                row=row_num,
                folder=str(row.get("folder") or "retry"),
                url=url,
            ))
    if not tasks:
        raise RuntimeError("No failed links found in report.")
    return tasks


def completed_urls_from_report(output):
    report = Path(output or DEFAULT_OUTPUT) / LOG_NAME
    completed = set()
    if not report.exists():
        return completed
    rows = rows_from_file(report)
    if not rows:
        return completed
    headers = [str(cell).lower().strip() for cell in rows[0]]
    for row_data in rows[1:]:
        row_data = list(row_data) + [""] * (len(headers) - len(row_data))
        row = dict(zip(headers, row_data))
        status = str(row.get("status") or "").lower()
        url = str(row.get("url") or "").strip()
        if (status == "ok" or status.startswith("skipped")) and url:
            completed.add(url.lower())
    return completed


def apply_folder_pattern(tasks, pattern):
    if pattern == "{folder}":
        return tasks
    updated = []
    for task in tasks:
        provider = provider_name(task.url) if task.url.startswith(("http://", "https://")) else "search"
        folder = pattern.replace("{row}", str(task.row)).replace("{folder}", task.folder).replace("{provider}", provider)
        updated.append(LinkTask(row=task.row, folder=folder, url=task.url))
    return updated


def tasks_from_payload(payload, resolve_search=False, log=None):
    source_mode = (payload.get("sourceMode") or "file").strip()
    input_path = (payload.get("inputPath") or "").strip().strip('"')
    if source_mode == "retry":
        tasks = retry_failed_tasks(payload.get("output") or DEFAULT_OUTPUT)
        return apply_folder_pattern(tasks, payload.get("folderPattern") or "{folder}")
    if source_mode == "manual":
        folder = (payload.get("manualFolder") or "manual_test").strip() or "manual_test"
        manual_text = payload.get("manualLinks") or ""
        links = re.findall(r"https?://[^\s,]+", manual_text)
        rows = [[folder, *links]]
    elif source_mode == "search":
        terms = search_terms_from_payload(payload)
        if not terms:
            raise RuntimeError("No style codes/search terms found.")
        if resolve_search:
            provider = payload.get("searchProvider") or "google"
            if provider == "custom":
                tasks = search_custom_tasks(terms, payload.get("searchTemplate") or "", log=log, quality_mode=payload.get("qualityMode", "best"))
            elif provider == "koton":
                tasks = search_tasks_from_terms(terms, log=log, quality_mode=payload.get("qualityMode", "best"))
            else:
                tasks = search_google_tasks(terms, log=log, quality_mode=payload.get("qualityMode", "best"))
            if not tasks:
                raise RuntimeError("Search completed, but no image links were found.")
            if payload.get("resumeFromReport"):
                completed = completed_urls_from_report(payload.get("output") or DEFAULT_OUTPUT)
                tasks = [task for task in tasks if task.url.strip().lower() not in completed]
                if not tasks:
                    raise RuntimeError("Nothing to resume. Search links are already marked OK in report.")
            return apply_folder_pattern(tasks, payload.get("folderPattern") or "{folder}")
        provider_key = payload.get("searchProvider") or "google"
        provider = {
            "custom": "custom-search",
            "koton": "koton-search",
            "google": "google-images",
        }.get(provider_key, "google-images")
        return [LinkTask(row=index, folder=term, url=f"{provider}://{urllib.parse.quote(term)}") for index, term in enumerate(terms, start=1)]
    elif input_path:
        rows = rows_from_file(input_path)
    else:
        rows = rows_from_pasted_text(payload.get("pasteText") or "")
    tasks = tasks_from_rows(rows)
    if not tasks:
        raise RuntimeError("No valid links found.")
    if payload.get("resumeFromReport"):
        completed = completed_urls_from_report(payload.get("output") or DEFAULT_OUTPUT)
        tasks = [task for task in tasks if task.url.strip().lower() not in completed]
        if not tasks:
            raise RuntimeError("Nothing to resume. All detected links are already marked OK in report.")
    return apply_folder_pattern(tasks, payload.get("folderPattern") or "{folder}")


def preview_from_payload(payload):
    source_mode = (payload.get("sourceMode") or "file").strip()
    tasks = tasks_from_payload(payload, resolve_search=False)
    providers = {}
    folders = set()
    sample = []
    for task in tasks:
        provider = task.url.split("://", 1)[0].replace("-", "_") if source_mode == "search" else provider_name(task.url)
        providers[provider] = providers.get(provider, 0) + 1
        folders.add(task.folder)
        if len(sample) < MAX_PREVIEW_SAMPLE:
            sample.append({
                "row": task.row,
                "folder": task.folder,
                "provider": provider,
                "url": task.url,
            })
    return {
        "total": len(tasks),
        "folders": len(folders),
        "providers": providers,
        "sample": sample,
    }


# ---------------------------------------------------------------------------
# Job lifecycle
# ---------------------------------------------------------------------------
def start_job(payload):
    with STATE.lock:
        if STATE.running:
            raise RuntimeError("Download already running.")

    output = Path((payload.get("output") or str(DEFAULT_OUTPUT)).strip().strip('"'))
    workers = min(50, max(1, int(payload.get("workers") or DEFAULT_WORKERS)))
    timeout = int(payload.get("timeout") or DEFAULT_TIMEOUT)
    # Validate queue early so failures surface before state reset.
    tasks_from_payload(payload, resolve_search=False)

    # Apply notification preference for this job.
    STATE.notify_enabled = bool(payload.get("notifyEnabled", EFFECTIVE_NOTIFICATIONS))

    output.mkdir(parents=True, exist_ok=True)
    report = output / LOG_NAME
    preserve_report = bool(payload.get("resumeFromReport")) or (payload.get("sourceMode") == "retry")
    if not preserve_report:
        report.unlink(missing_ok=True)
    STATE.reset()
    STATE.update(
        running=True,
        output=str(output.resolve()),
        report=str(report.resolve()),
        status="Preparing queue",
        notify_enabled=STATE.notify_enabled,
    )
    STATE.mark_started()
    STATE.log(f"Output: {output.resolve()}")
    removed_metadata = cleanup_output_metadata(output)
    if removed_metadata:
        STATE.log(f"Cleaned {removed_metadata} metadata file(s) from output.")
    proxy = payload.get("proxy") or {}
    speed_limit = int(payload.get("speedLimitKbps") or DEFAULT_SPEED_LIMIT_KBPS)
    if proxy and (proxy.get("http") or proxy.get("https")):
        STATE.log(f"Proxy enabled: {proxy}")
    if speed_limit:
        STATE.log(f"Speed limit: {speed_limit} KB/s")

    thread = threading.Thread(
        target=prepare_and_run,
        args=(payload, output, workers, timeout, STATE.stop_event, STATE.pause_event),
        daemon=True,
    )
    thread.start()


def prepare_and_run(payload, output, workers, timeout, stop_event, pause_event):
    try:
        tasks = tasks_from_payload(payload, resolve_search=True, log=STATE.log)
        providers = {}
        for task in tasks:
            providers[provider_name(task.url)] = providers.get(provider_name(task.url), 0) + 1
        STATE.update(
            total=len(tasks),
            status=f"Running {len(tasks)} link(s)",
            providers=providers,
        )
        STATE.log(f"Started: {len(tasks)} link(s)")
        run_download(tasks, output, workers, timeout, stop_event, pause_event, payload)
    except Exception as exc:
        with STATE.lock:
            STATE.running = False
            STATE.status = "Error"
        STATE.log(f"ERROR: {exc}")


def run_download(tasks, output, workers, timeout, stop_event, pause_event, payload):
    failed_for_browser = []

    def log(text):
        STATE.log(text)

    def transfer_progress(byte_count, url, task=None):
        STATE.add_bytes(byte_count, url, task)

    def record_result(task, result, done=None, total=None, is_retry=False):
        target = Path(output) / sanitize_name(task.folder)
        failure_group = classify_failure(result.status, result.message)
        with STATE.lock:
            if done is not None:
                STATE.done = done
            if result.status == "ok":
                STATE.ok += 1
            elif result.status.startswith("skipped"):
                STATE.skipped += 1
            elif not is_retry:
                STATE.failed += 1
            STATE.files += result.files
            if failure_group and failure_group != "skipped":
                STATE.failure_groups[failure_group] = STATE.failure_groups.get(failure_group, 0) + 1
            if result.status != "ok" and hasattr(result, "attempts") and result.attempts > 1:
                STATE.retry_queue[failure_group or result.status] = STATE.retry_queue.get(failure_group or result.status, 0) + 1
            if done is not None and total is not None:
                remaining = max(0, total - done)
                elapsed = max(1, time.time() - STATE.start_time) if STATE.start_time else 1
                eta = int(remaining / (done / elapsed)) if done else 0
                STATE.status = f"Progress {done}/{total} | ETA {format_seconds(eta)}"
        STATE.add_row({
            "row": task.row,
            "folder": task.folder,
            "status": result.status,
            "files": result.files,
            "message": result.message,
            "attempts": getattr(result, "attempts", 1),
            "duration": getattr(result, "duration", 0),
            "bytes": getattr(result, "bytes_downloaded", 0),
            "thumb": first_thumbnail(target),
            "failure": failure_group,
        })
        retry_note = f" | attempts {getattr(result, 'attempts', 1)}" if getattr(result, "attempts", 1) > 1 else ""
        STATE.log(f"[{task.row}] {task.folder} -> {result.status} | {result.files} file(s){retry_note} | {result.message}")

    def progress(done, total, task, result):
        needs_browser = result.status in ("login_required", "html_page")
        if (payload.get("browserMode") or needs_browser) and result.status != "ok" and not result.status.startswith("skipped"):
            failed_for_browser.append(task)
        record_result(task, result, done=done, total=total)

    try:
        if payload.get("browserMode"):
            STATE.log("Browser fallback selected. Direct download runs first; failed links retry in a hidden browser session.")
        downloader = UniversalDownloader(
            output,
            workers=workers,
            timeout=timeout,
            log=log,
            skip_existing=payload.get("skipExisting", False),
            file_mode=payload.get("fileMode", "all"),
            duplicate_detection=payload.get("dedupe", False),
            max_retries=2,
            provider_limits={
                key: min(value, workers) for key, value in PROVIDER_LIMITS.items()
            },
            transfer_progress=transfer_progress,
            task_events=STATE.task_event,
            proxy=payload.get("proxy"),
            speed_limit_kbps=int(payload.get("speedLimitKbps") or 0)
        )
        if payload.get("sourceMode") == "retry" or payload.get("resumeFromReport"):
            report_path = Path(output) / LOG_NAME
            existing_rows = rows_from_file(report_path)
            if existing_rows and len(existing_rows) > 1:
                downloader.report_rows = list(existing_rows[1:])
        
        with STATE.lock:
            STATE.downloader = downloader

        downloader.run(tasks, progress=progress, stop_event=stop_event, pause_event=pause_event)
        if failed_for_browser and not stop_event.is_set():
            STATE.log(f"Browser fallback pass: {len(failed_for_browser)} failed link(s)")
            before_failed = STATE.failed

            def browser_progress(task, result):
                if result.status == "ok":
                    with STATE.lock:
                        STATE.failed = max(0, STATE.failed - 1)
                record_result(task, result, is_retry=True)
                downloader.append_report(task, result)

            browser_fallback_download(failed_for_browser, output, browser_progress)
            downloader.write_xlsx_report()
            STATE.log(f"Browser fallback finished. Previous failed count before fallback: {before_failed}")
        with STATE.lock:
            STATE.running = False
            STATE.downloader = None
            STATE.status = "Stopped" if stop_event.is_set() else f"Done: {STATE.ok} ok, {STATE.skipped} skipped, {STATE.failed} failed"
            cleanup_output_metadata(output)
            STATE.audit = output_audit(output, tasks)
            notify = STATE.notify_enabled
            ok_count = STATE.ok
            failed_count = STATE.failed
            files_count = STATE.files
        STATE.log("Stopped by user." if stop_event.is_set() else "Finished.")
        if notify and not stop_event.is_set():
            _send_desktop_notification(
                "Media Download Manager",
                f"Done: {ok_count} ok, {STATE.skipped} skipped, {failed_count} failed, {files_count} files",
            )
    except Exception as exc:
        with STATE.lock:
            STATE.running = False
            STATE.status = "Error"
        STATE.log(f"ERROR: {exc}")


def stop_job():
    with STATE.lock:
        if not STATE.running:
            return
        STATE.stop_event.set()
        STATE.status = "Stopping..."
    STATE.log("Stop requested. Running downloads will finish or cancel best effort.")


def pause_job():
    message = ""
    with STATE.lock:
        if not STATE.running:
            return
        if STATE.pause_event.is_set():
            STATE.pause_event.clear()
            STATE.status = f"Progress {STATE.done}/{STATE.total}"
            message = "Resumed."
        else:
            STATE.pause_event.set()
            STATE.status = "Paused"
            message = "Paused. Running downloads may finish; new downloads will wait."
    STATE.log(message)


# ---------------------------------------------------------------------------
# Saved job configs (load/save/delete)
# ---------------------------------------------------------------------------
def _safe_config_name(name: str) -> str:
    name = sanitize_name(name or "default")
    return name or "default"


def save_job_config(name: str, settings: dict) -> Path:
    name = _safe_config_name(name)
    path = SAVED_CONFIGS_DIR / f"{name}.json"
    with path.open("w", encoding="utf-8") as handle:
        json.dump({"name": name, "settings": settings}, handle, indent=2, ensure_ascii=False)
    return path


def load_job_config(name: str) -> dict:
    name = _safe_config_name(name)
    path = SAVED_CONFIGS_DIR / f"{name}.json"
    if not path.exists():
        raise RuntimeError(f"No saved config named '{name}'.")
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    return data.get("settings", {})


def delete_job_config(name: str) -> None:
    name = _safe_config_name(name)
    path = SAVED_CONFIGS_DIR / f"{name}.json"
    if path.exists():
        path.unlink()
    else:
        raise RuntimeError(f"No saved config named '{name}'.")


# ---------------------------------------------------------------------------
# HTTP handler
# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):

    # ------------------------------------------------------------------
    # Common helpers
    # ------------------------------------------------------------------
    def _client_ip(self) -> str:
        return self.client_address[0]

    def _add_security_headers(self):
        """Add secure HTTP headers to every response."""
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "SAMEORIGIN")
        self.send_header("Referrer-Policy", "strict-origin-when-cross-origin")
        self.send_header(
            "Content-Security-Policy",
            "default-src 'self'; "
            "script-src 'self' 'unsafe-inline' https://unpkg.com https://cdn.jsdelivr.net; "
            "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
            "font-src 'self' https://fonts.gstatic.com; "
            "img-src 'self' data: blob:; "
            "connect-src 'self';"
        )

    def _add_cors_headers(self):
        """Add CORS headers for allowed origins."""
        origin = self.headers.get("Origin", "")
        if ALLOWED_ORIGINS and origin in ALLOWED_ORIGINS:
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")
        elif not ALLOWED_ORIGINS:
            # No restriction configured — allow same-origin only (browser handles it)
            pass

    def _redirect(self, location: str, status: int = 302):
        self.send_response(status)
        self.send_header("Location", location)
        self._add_security_headers()
        self.end_headers()

    def _require_auth(self) -> bool:
        """Return True if request is authenticated. Otherwise redirect and return False."""
        if _is_authenticated(self):
            return True
        if self.path.startswith("/api/"):
            self.send_json({"error": "Unauthorized"}, status=401)
        else:
            self._redirect("/login")
        return False

    # ------------------------------------------------------------------
    # GET
    # ------------------------------------------------------------------
    def do_GET(self):
        path = urllib.parse.urlsplit(self.path).path

        # Rate limiting
        if not _check_rate_limit(self._client_ip()):
            self.send_response(429)
            self.send_header("Retry-After", str(RATE_LIMIT_WINDOW))
            self._add_security_headers()
            self.end_headers()
            return

        try:
            if path == "/login":
                self.send_html(_load_login_html())
                return
            if path == "/auth/logout":
                self.send_response(302)
                self.send_header("Location", "/login")
                self.send_header(
                    "Set-Cookie",
                    f"{SESSION_COOKIE_NAME}=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0"
                )
                self._add_security_headers()
                self.end_headers()
                return
            if not self._require_auth():
                return
            if path == "/":
                page = _load_html().replace("__DEFAULT_OUTPUT__", str(DEFAULT_OUTPUT))
                page = page.replace("__DEFAULT_WORKERS__", str(EFFECTIVE_WORKERS))
                page = page.replace("__DEFAULT_TIMEOUT__", str(DEFAULT_TIMEOUT))
                self.send_html(page)
            elif path == "/api/status":
                self.send_json(STATE.snapshot())
            elif path == "/api/select-output":
                self.send_json({"path": pick_output()})
            elif path == "/api/thumb":
                query = urllib.parse.parse_qs(urllib.parse.urlsplit(self.path).query)
                thumb_path = query.get("path", [""])[0]
                self.serve_thumbnail(thumb_path)
            else:
                self.send_error(404)
        except Exception as exc:
            logger.exception("GET error: %s", exc)
            self.send_error(500, str(exc))

    # ------------------------------------------------------------------
    # POST
    # ------------------------------------------------------------------
    def do_POST(self):
        path = urllib.parse.urlsplit(self.path).path

        # Rate limiting
        if not _check_rate_limit(self._client_ip()):
            self.send_response(429)
            self.send_header("Retry-After", str(RATE_LIMIT_WINDOW))
            self._add_security_headers()
            self.end_headers()
            return

        # Auth login endpoint — no session required
        if path == "/auth/login":
            self._handle_login()
            return

        # All other POST endpoints require authentication
        if not self._require_auth():
            return

        try:
            if path == "/api/upload":
                # Handle raw file upload
                length = int(self.headers.get("Content-Length", "0"))
                filename = self.headers.get("X-Filename", "uploaded_file.csv")
                safe_name = "".join(c for c in filename if c.isalnum() or c in ".-_")
                if not safe_name: safe_name = "upload.csv"
                file_path = UPLOAD_DIR / f"{int(time.time())}_{safe_name}"
                with open(file_path, "wb") as f:
                    # Stream the upload to avoid loading large files in memory
                    remaining = length
                    while remaining > 0:
                        chunk = self.rfile.read(min(remaining, 65536))
                        if not chunk: break
                        f.write(chunk)
                        remaining -= len(chunk)
                self.send_json({"path": str(file_path.resolve())})
                return

            payload = self.read_json()
            if path == "/api/start":
                start_job(payload)
                self.send_json({"status": "started"})
            elif path == "/api/enqueue":
                if not STATE.running or not STATE.downloader:
                    # Fallback: start a new job if nothing is running
                    payload["sourceMode"] = "manual"
                    start_job(payload)
                    self.send_json({"status": "started"})
                else:
                    new_tasks = tasks_from_payload(payload, resolve_search=True, log=STATE.log)
                    if new_tasks:
                        for t in new_tasks:
                            STATE.downloader.enqueue_task(t)
                        with STATE.lock:
                            STATE.total += len(new_tasks)
                            for t in new_tasks:
                                STATE.providers[provider_name(t.url)] = STATE.providers.get(provider_name(t.url), 0) + 1
                        STATE.log(f"Enqueued {len(new_tasks)} new link(s).")
                    self.send_json({"status": "enqueued", "added": len(new_tasks)})
            elif path == "/api/preview":
                self.send_json(preview_from_payload(payload))
            elif path == "/api/stop":
                stop_job()
                self.send_json({"ok": True})
            elif path == "/api/pause":
                pause_job()
                self.send_json({"ok": True})
            elif path == "/api/open-output":
                open_path(payload.get("path") or DEFAULT_OUTPUT)
                self.send_json({"ok": True})
            elif path == "/api/open-report":
                report = Path(payload.get("path") or DEFAULT_OUTPUT) / LOG_NAME
                if report.exists():
                    open_in_file_manager(report.resolve())
                else:
                    raise RuntimeError("Report file is not available yet.")
                self.send_json({"ok": True})
            elif path == "/api/create-template":
                create_excel_template(payload.get("path") or DEFAULT_OUTPUT)
                self.send_json({"ok": True})
            elif path == "/api/rename-files":
                self.send_json(rename_downloaded_files(payload.get("path") or DEFAULT_OUTPUT))
            elif path == "/api/save-config":
                save_job_config(payload.get("name", "default"), payload.get("settings", {}))
                self.send_json({"ok": True})
            elif path == "/api/load-config":
                self.send_json({"settings": load_job_config(payload.get("name", "default"))})
            elif path == "/api/delete-config":
                delete_job_config(payload.get("name", "default"))
                self.send_json({"ok": True})
            else:
                self.send_json({"error": "API endpoint not found."}, status=404)
        except Exception as exc:
            logger.exception("POST error: %s", exc)
            self.send_json({"error": str(exc)}, status=400)

    def do_OPTIONS(self):
        """Handle CORS preflight."""
        self.send_response(204)
        self._add_cors_headers()
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self._add_security_headers()
        self.end_headers()

    # ------------------------------------------------------------------
    # Auth login handler
    # ------------------------------------------------------------------
    def _handle_login(self):
        if not AUTH_ENABLED:
            self._redirect("/")
            return
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length).decode("utf-8", errors="replace") if length else ""
        params = urllib.parse.parse_qs(body)
        username = params.get("username", [""])[0].strip()
        password = params.get("password", [""])[0]
        if username == AUTH_USERNAME and password == AUTH_PASSWORD:
            token = _make_session_token(username)
            self.send_response(302)
            self.send_header("Location", "/")
            self.send_header(
                "Set-Cookie",
                f"{SESSION_COOKIE_NAME}={token}; Path=/; HttpOnly; SameSite=Strict; Max-Age={SESSION_MAX_AGE}"
            )
            self._add_security_headers()
            self.end_headers()
        else:
            self.send_html(_load_login_html(error="Invalid username or password."))

    # ------------------------------------------------------------------
    # Response helpers
    # ------------------------------------------------------------------
    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length > 2 * 1024 * 1024:  # 2 MB max payload
            raise ValueError("Request body too large.")
        raw = self.rfile.read(length).decode("utf-8-sig") if length else "{}"
        return json.loads(raw or "{}")

    def send_json(self, data, status=200):
        raw = json.dumps(data).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self._add_security_headers()
        self._add_cors_headers()
        self.end_headers()
        self.wfile.write(raw)

    def send_html(self, text, status=200):
        raw = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self._add_security_headers()
        self.end_headers()
        self.wfile.write(raw)

    def serve_thumbnail(self, path):
        """Serve a thumbnail SAFELY.

        - Path must resolve under the output directory to block traversal.
        - Only image extensions are served.
        - Files above MAX_THUMB_SIZE are rejected.
        - Streamed in chunks to avoid loading large files into memory.
        """
        if not path:
            self.send_error(404)
            return
        file_path = Path(path).expanduser()
        try:
            resolved = file_path.resolve()
        except (OSError, RuntimeError):
            self.send_error(404)
            return

        # Lock thumbnails to the output directory tree only.
        allowed_root = Path(str(STATE.output)).resolve()
        try:
            resolved.relative_to(allowed_root)
        except ValueError:
            self.send_error(403)
            return

        if not resolved.is_file():
            self.send_error(404)
            return
        if resolved.suffix.lower() not in THUMB_EXTENSIONS:
            self.send_error(403)
            return
        if resolved.stat().st_size > MAX_THUMB_SIZE:
            self.send_error(413)
            return

        content_type = mimetypes.guess_type(str(resolved))[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(resolved.stat().st_size))
        self._add_security_headers()
        self.end_headers()
        with resolved.open("rb") as handle:
            while True:
                chunk = handle.read(1024 * 256)
                if not chunk:
                    break
                self.wfile.write(chunk)

    def log_message(self, fmt, *args):
        if PRODUCTION:
            logger.info(fmt % args)


def create_server():
    last_error = None
    for port in range(PORT, PORT + PORT_FALLBACK_LIMIT + 1):
        try:
            return ThreadingHTTPServer((HOST, port), Handler), port
        except OSError as exc:
            last_error = exc
            if getattr(exc, "winerror", None) not in {10013, 10048}:
                raise
    try:
        server = ThreadingHTTPServer((HOST, 0), Handler)
        return server, server.server_address[1]
    except OSError:
        if last_error:
            raise last_error
        raise


def _get_lan_ip() -> str:
    """Best-effort local LAN IP detection."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return ""


def main():
    server, port = create_server()
    local_url = f"http://127.0.0.1:{port}/"
    lan_ip = _get_lan_ip()
    lan_url = f"http://{lan_ip}:{port}/" if lan_ip and lan_ip != "127.0.0.1" else ""

    print(f"\n{'='*55}")
    print(f"  Media Download Manager  v{APP_VERSION}")
    print(f"{'='*55}")
    print(f"  Local:   {local_url}")
    if lan_url:
        print(f"  LAN:     {lan_url}")
    if TUNNEL_URL:
        print(f"  Tunnel:  {TUNNEL_URL}")
    if AUTH_ENABLED:
        print(f"  Auth:    ON  (username: {AUTH_USERNAME})")
    else:
        print(f"  Auth:    OFF (open access)")
    print(f"{'='*55}\n")

    logger.info("Server started on %s:%s", HOST, port)

    # Auto-open browser (skip on headless/Docker)
    if not os.environ.get("NO_BROWSER"):
        webbrowser.open(local_url)

    server.serve_forever()


if __name__ == "__main__":
    main()
