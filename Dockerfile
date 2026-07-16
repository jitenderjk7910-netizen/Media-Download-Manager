# ---------------------------------------------------------------------------
# Media Download Manager - Dockerfile
# ---------------------------------------------------------------------------
FROM python:3.11-slim

# Security: run as non-root
RUN addgroup --system mdm && adduser --system --ingroup mdm mdm

WORKDIR /app

# Install system deps for Playwright (optional browser fallback)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 libnss3 libatk1.0-0 libatk-bridge2.0-0 \
    libcups2 libdbus-1-3 libdrm2 libgbm1 libgtk-3-0 \
    libx11-6 libxcomposite1 libxdamage1 libxext6 libxfixes3 \
    libxrandr2 libxshmfence1 && \
    rm -rf /var/lib/apt/lists/*

# Install Python dependencies first (layer cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Create runtime directories and set ownership
RUN mkdir -p downloads uploads templates/saved_configs && \
    chown -R mdm:mdm /app

USER mdm

EXPOSE 8080

# NO_BROWSER: skip webbrowser.open() in headless Docker
ENV NO_BROWSER=1 \
    PRODUCTION=1 \
    HOST=0.0.0.0 \
    PORT=8080

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8080/api/status')"

CMD ["python", "web_app.py"]
