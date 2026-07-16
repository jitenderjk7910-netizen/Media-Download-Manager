"""Centralized configuration for Media Download Manager.

All hardcoded values (timeouts, limits, semaphores, paths) live here so the rest
of the codebase imports named constants instead of using magic numbers. A user
can optionally override any value by creating a ``config.json`` next to this
file (see ``config.example.json`` for the full list of keys).
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, Optional

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
APP_DIR: Path = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

DEFAULT_OUTPUT: Path = APP_DIR / "downloads"
UPLOAD_DIR: Path = APP_DIR / "uploads"
TEMPLATES_DIR: Path = APP_DIR / "templates"
SAVED_CONFIGS_DIR: Path = TEMPLATES_DIR / "saved_configs"
LOG_NAME: str = "download_report.xlsx"
CONFIG_FILE: Path = APP_DIR / "config.json"

# Persistent browser profile used by the Playwright fallback (logged-in session).
AUTH_PROFILE: Path = Path(
    os.environ.get("LOCALAPPDATA", str(APP_DIR))
) / "UniversalLinkDownloader" / "browser_profile"

# ---------------------------------------------------------------------------
# App metadata
# ---------------------------------------------------------------------------
APP_TITLE: str = "Media Download Manager"
APP_VERSION: str = "2.0"

# ---------------------------------------------------------------------------
# Web server
# ---------------------------------------------------------------------------
HOST: str = os.environ.get("HOST", "0.0.0.0")
PORT: int = int(os.environ.get("PORT", "8080"))
PORT_FALLBACK_LIMIT: int = 40

# ---------------------------------------------------------------------------
# Built-in authentication (optional — leave blank to run open)
# ---------------------------------------------------------------------------
AUTH_USERNAME: str = os.environ.get("USERNAME", "")
AUTH_PASSWORD: str = os.environ.get("PASSWORD", "")
# Secret used to sign session cookies. Auto-generated if not provided.
SESSION_SECRET: str = os.environ.get(
    "SESSION_SECRET",
    __import__("secrets").token_hex(32),
)
SESSION_COOKIE_NAME: str = "mdm_session"
SESSION_MAX_AGE: int = 86400  # 24 hours

# ---------------------------------------------------------------------------
# Tunnel / reverse-proxy shareable URL
# ---------------------------------------------------------------------------
TUNNEL_URL: str = os.environ.get("TUNNEL_URL", "")

# ---------------------------------------------------------------------------
# CORS
# ---------------------------------------------------------------------------
ALLOWED_ORIGINS: list = [
    o.strip()
    for o in os.environ.get("ALLOWED_ORIGINS", "").split(",")
    if o.strip()
]

# ---------------------------------------------------------------------------
# Rate limiting
# ---------------------------------------------------------------------------
RATE_LIMIT_REQUESTS: int = int(os.environ.get("RATE_LIMIT_REQUESTS", "120"))
RATE_LIMIT_WINDOW: int = int(os.environ.get("RATE_LIMIT_WINDOW", "60"))  # seconds

# ---------------------------------------------------------------------------
# Production mode
# ---------------------------------------------------------------------------
PRODUCTION: bool = os.environ.get("PRODUCTION", "0").strip() in ("1", "true", "yes")


# ---------------------------------------------------------------------------
# Download defaults
# ---------------------------------------------------------------------------
DEFAULT_WORKERS: int = 50
MIN_WORKERS: int = 1
MAX_WORKERS: int = 50
DEFAULT_TIMEOUT: int = 60
MIN_TIMEOUT: int = 10
MAX_TIMEOUT: int = 180
DEFAULT_MAX_RETRIES: int = 2
MAX_RETRIES_LIMIT: int = 5
PAUSE_POLL_INTERVAL: float = 0.2  # seconds between pause-event checks
RETRY_BACKOFF_MAX: int = 4  # seconds

# Provider concurrency caps (BoundedSemaphore per provider)
PROVIDER_LIMITS: Dict[str, int] = {
    "sharepoint_onedrive": 10,
    "google_drive": 8,
    "dropbox": 10,
    "imgbb": 20,
    "generic": MAX_WORKERS,
}

# ---------------------------------------------------------------------------
# File / parsing limits (named constants replacing magic numbers)
# ---------------------------------------------------------------------------
MAX_PAGE_IMAGES: int = 80        # generic web page image extraction
MAX_ONEDRIVE_IMAGES: int = 30    # OneDrive viewer page
MAX_BROWSER_IMAGES: int = 40     # Playwright visible-image scan
MAX_PREVIEW_SAMPLE: int = 80     # preview payload sample rows
MAX_DISPLAYED_ROWS: int = 300    # classic GUI load preview
MAX_REPORT_ROWS: int = 500       # status snapshot rows kept in memory
MAX_LOG_LINES: int = 600         # status snapshot logs kept in memory
MIN_BROWSER_IMAGE_AREA: int = 10000  # px^2, ignore tiny images in screenshot fallback

# ---------------------------------------------------------------------------
# HTTP transport
# ---------------------------------------------------------------------------
CHUNK_SIZE: int = 1024 * 256  # 256 KB streaming chunks
PARTIAL_DIR_NAME: str = ".partials"
MAGIC_BYTE_BUFFER: int = 16   # extension fix
VALIDATION_BUFFER: int = 512  # payload signature validation
MAX_THUMB_SIZE: int = 10 * 1024 * 1024  # 10 MB cap for /api/thumb serving

DEFAULT_USER_AGENT: str = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
DEFAULT_SESSION_HEADERS: dict = {
    "User-Agent": DEFAULT_USER_AGENT,
    "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Sec-Fetch-Dest": "image",
    "Sec-Fetch-Mode": "no-cors",
    "Sec-Fetch-Site": "cross-site",
    "Connection": "keep-alive",
}
DROPBOX_USER_AGENT: str = "python-requests/2.32"

# File extension groups
IMAGE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".webp", ".gif",
    ".bmp", ".tif", ".tiff", ".heic", ".heif",
}
VIDEO_EXTENSIONS = {".mp4", ".mov", ".m4v", ".avi", ".mkv", ".webm"}
ARCHIVE_EXTENSIONS = {".zip"}
THUMB_EXTENSIONS = IMAGE_EXTENSIONS  # only these can be served by /api/thumb

# Retryable HTTP status codes / message fragments
RETRYABLE_STATUSES = {"timeout", "network_error", "server_error", "rate_limited"}
RETRYABLE_MESSAGE_FRAGMENTS = (
    "timed out", "timeout", "connection reset", "connection aborted",
    "temporarily unavailable", "remote end closed",
    "429", "500", "502", "503", "504",
)

# ---------------------------------------------------------------------------
# Browser fallback timeouts (Playwright / Selenium)
# ---------------------------------------------------------------------------
BROWSER_PAGE_TIMEOUT_MS: int = 20_000
BROWSER_CONTROL_TIMEOUT_MS: int = 1_000
BROWSER_CLICK_TIMEOUT_MS: int = 5_000
BROWSER_DOWNLOAD_TIMEOUT_MS: int = 30_000
BROWSER_IMAGE_CHECK_TIMEOUT_MS: int = 500
BROWSER_SCREENSHOT_TIMEOUT_MS: int = 10_000

# ---------------------------------------------------------------------------
# New feature defaults (proxy / speed limiting / notifications)
# ---------------------------------------------------------------------------
DEFAULT_PROXY: Dict[str, str] = {"http": "", "https": ""}
DEFAULT_SPEED_LIMIT_KBPS: int = 0  # 0 = unlimited
DEFAULT_NOTIFICATIONS_ENABLED: bool = True
# Short aliases used by the web layer.
DEFAULT_NOTIFICATIONS: bool = DEFAULT_NOTIFICATIONS_ENABLED


# ---------------------------------------------------------------------------
# Config loader
# ---------------------------------------------------------------------------
def _coalesce(value: Any, fallback: Any) -> Any:
    """Return ``value`` if it is not None, else ``fallback``."""
    return value if value is not None else fallback


def load_config(path: Optional[Path] = None) -> Dict[str, Any]:
    """Load optional JSON overrides from ``config.json``.

    Missing keys fall back to the module-level defaults. A missing or malformed
    file is never fatal — we just warn and continue with defaults.
    """
    path = path or CONFIG_FILE
    overrides: Dict[str, Any] = {}
    if path.exists():
        try:
            with path.open("r", encoding="utf-8") as handle:
                loaded = json.load(handle)
            if isinstance(loaded, dict):
                overrides = loaded
        except (OSError, json.JSONDecodeError) as exc:
            print(f"[app_config] Could not read {path.name}: {exc}; using defaults.")
    return overrides


# User overrides loaded once at import time (safe — edit config.json + restart).
USER_CONFIG: Dict[str, Any] = load_config()

# Effective values (defaults overridden by config.json where present).
EFFECTIVE_HOST: str = str(_coalesce(USER_CONFIG.get("host"), HOST))
EFFECTIVE_PORT: int = int(_coalesce(USER_CONFIG.get("port"), PORT))
EFFECTIVE_WORKERS: int = int(_coalesce(USER_CONFIG.get("workers"), DEFAULT_WORKERS))
EFFECTIVE_TIMEOUT: int = int(_coalesce(USER_CONFIG.get("timeout"), DEFAULT_TIMEOUT))
EFFECTIVE_PROXY: Dict[str, str] = {
    "http": str((_coalesce(USER_CONFIG.get("proxy"), DEFAULT_PROXY) or {}).get("http", "")),
    "https": str((_coalesce(USER_CONFIG.get("proxy"), DEFAULT_PROXY) or {}).get("https", "")),
}
EFFECTIVE_SPEED_LIMIT_KBPS: int = int(
    _coalesce(USER_CONFIG.get("speed_limit_kbps"), DEFAULT_SPEED_LIMIT_KBPS)
)
EFFECTIVE_NOTIFICATIONS: bool = bool(
    _coalesce(USER_CONFIG.get("enable_notifications"), DEFAULT_NOTIFICATIONS_ENABLED)
)
