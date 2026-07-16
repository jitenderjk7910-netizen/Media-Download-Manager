# Media Download Manager

A production-ready, secure web-based tool for batch downloading images from direct URLs, Dropbox, Google Drive, SharePoint, OneDrive, ImgBB, and generic web pages.

## Features

- **Batch Download** from Excel/CSV files
- **Manual Mode** — paste links directly
- **Search & Download** — search product images by style codes
- **Secure by default** — optional username/password login, secure HTTP headers, rate limiting
- **LAN sharing** — colleagues connect via browser, no installation needed
- **Docker ready** — one command to deploy
- **Tunnel support** — share securely via Cloudflare Tunnel, Tailscale Funnel, or ngrok

---

## Quick Start (Local)

### 1. Install requirements
```bat
pip install -r requirements.txt
```

### 2. Run (no authentication)
```bat
python web_app.py
```
The app opens automatically at `http://localhost:8080/` and also displays your LAN address.

### 3. Run (with authentication)
Set environment variables before running:
```powershell
$env:MDM_USERNAME="admin"; $env:MDM_PASSWORD="yourpassword"; python web_app.py
```

Or create a `.env` file (copy from `.env.example`):
```
MDM_USERNAME=admin
MDM_PASSWORD=yourpassword
```
Then load it:
```powershell
Get-Content .env | ForEach-Object { if ($_ -match "^([^#][^=]+)=(.*)$") { [System.Environment]::SetEnvironmentVariable($Matches[1], $Matches[2]) } }
python web_app.py
```

---

## Running with Docker

### 1. Copy and configure environment
```bash
cp .env.example .env
# Edit .env with your MDM_USERNAME, MDM_PASSWORD, etc.
```

### 2. Build and start
```bash
docker-compose up -d
```

### 3. Open in browser
```
http://localhost:8080/
```

### Stop
```bash
docker-compose down
```

Downloaded files persist in `./downloads/` (mounted as a volume).

---

## Sharing over LAN (Office Wi-Fi)

When you run `python web_app.py`, the startup banner displays your LAN address automatically:

```
=======================================================
  Media Download Manager  v2.0
=======================================================
  Local:   http://127.0.0.1:8080/
  LAN:     http://192.168.1.65:8080/
  Auth:    OFF (open access)
=======================================================
```

Your colleagues on the **same Wi-Fi network** can open `http://192.168.1.65:8080/` in their browser and use the tool. They never see your source code or filesystem.

> **Tip:** Keep the terminal window open while they use it. Closing it stops the server.

---

## Sharing Securely over the Internet

### Option A: Cloudflare Tunnel (Recommended — Free)

1. Install: https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/
2. Run your app normally: `python web_app.py`
3. In another terminal:
   ```bash
   cloudflared tunnel --url http://localhost:8080
   ```
4. Copy the `trycloudflare.com` URL shown and set it in `.env`:
   ```
   TUNNEL_URL=https://your-tunnel.trycloudflare.com
   ```
5. Restart the app — the tunnel URL will show in the startup banner and users can access it from anywhere.

> Always enable `MDM_USERNAME` and `MDM_PASSWORD` when sharing over the internet!

### Option B: Tailscale Funnel

1. Install Tailscale: https://tailscale.com/download
2. Enable Funnel:
   ```bash
   tailscale funnel 8080
   ```
3. Your app is now accessible at `https://yourdevice.tailXXXXX.ts.net/`
4. Set `TUNNEL_URL=https://yourdevice.tailXXXXX.ts.net/` in `.env`

### Option C: ngrok

1. Install ngrok: https://ngrok.com/download
2. Run your app: `python web_app.py`
3. In another terminal:
   ```bash
   ngrok http 8080
   ```
4. Copy the `https://xxxxx.ngrok-free.app` URL shown.

---

## Configuration Reference

All settings can be configured via environment variables or a `.env` file.

| Variable | Default | Description |
|---|---|---|
| `HOST` | `0.0.0.0` | Server bind address |
| `PORT` | `8080` | Server port |
| `MDM_USERNAME` | *(blank)* | Login username (leave blank for open access) |
| `MDM_PASSWORD` | *(blank)* | Login password |
| `SESSION_SECRET` | *(auto)* | Cookie signing secret (set for production) |
| `TUNNEL_URL` | *(blank)* | Public URL to display on startup |
| `ALLOWED_ORIGINS` | *(blank)* | Comma-separated CORS origins |
| `RATE_LIMIT_REQUESTS` | `120` | Max requests per IP per window |
| `RATE_LIMIT_WINDOW` | `60` | Rate limit window in seconds |
| `PRODUCTION` | `0` | Set `1` to enable JSON logging to `app.log` |
| `NO_BROWSER` | `0` | Set `1` to skip auto-opening browser |

---

## Input Format

```text
BarcodeOrFolder    Link1                          Link2
ABC001             https://example.com/image.jpg  https://drive.google.com/...
ABC002             https://dropbox.com/...         https://ibb.co/...
```

The first non-link column becomes the output folder name.

---

## Security Model

| Feature | Detail |
|---|---|
| Authentication | HMAC-signed session cookie (24h), optional via env vars |
| Secure Headers | X-Frame-Options, X-Content-Type-Options, CSP, Referrer-Policy |
| Rate Limiting | 120 requests/min per IP (configurable) |
| CORS | Same-origin by default; opt-in via `ALLOWED_ORIGINS` |
| File Access | Downloads sandboxed to `downloads/` folder |
| Directory Traversal | Blocked — thumbnails locked to output directory |
| Docker | Runs as non-root `mdm` user |

---

## Supported Sources

- Direct image/file URLs (`.jpg`, `.png`, `.webp`, `.gif`, `.mp4`, `.zip`, etc.)
- Dropbox shared folder links
- Google Drive files and folders (via `gdown`)
- SharePoint / OneDrive shared links
- ImgBB and generic image pages
- Browser fallback (Playwright) for sites requiring login
