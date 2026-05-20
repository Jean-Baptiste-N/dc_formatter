# DC Formatter - All-in-One Image
# Single container with API, Frontend, and Pipeline

FROM python:3.11-slim

# Set environment variables for Python
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# Set working directory
WORKDIR /app

# Install system dependencies required for lxml and other packages
RUN apt-get update && apt-get install -y --no-install-recommends \
    libxml2-dev \
    libxslt-dev \
    build-essential \
    gosu \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements (API + pipeline)
COPY requirements-api.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements-api.txt

# Copy application code tools and template
COPY tools /app/tools
COPY TEMPLATE /app/TEMPLATE
COPY app.py /app/app.py
COPY index.html /app/index.html
COPY entrypoint.sh /app/entrypoint.sh

# Create non-root user for security with home directory
RUN groupadd -g 1000 dcformatter && \
    useradd -m -u 1000 -g 1000 -s /bin/bash -c "Docker app user" dcformatter && \
    chown -R dcformatter:dcformatter /app && \
    mkdir -p /home/dcformatter/.vscode-server && \
    chown -R dcformatter:dcformatter /home/dcformatter && \
    chmod +x /app/entrypoint.sh

# Expose API port
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:8000/health')" || exit 1

# Use entrypoint script to fix permissions and run as dcformatter
ENTRYPOINT ["/app/entrypoint.sh"]
