# Build stage for Python 3.11
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
    && rm -rf /var/lib/apt/lists/*

# Copy production requirements first for better layer caching
COPY requirements-prod.txt .

# Install Python dependencies (production-only)
RUN pip install --no-cache-dir -r requirements-prod.txt

# Copy application code (tools3 module and assets)
COPY tools3 /app/tools3
COPY assets /app/assets

# Create non-root user for security
RUN groupadd -g 1000 dcformatter && \
    useradd -u 1000 -g 1000 -s /sbin/nologin -c "Docker app user" dcformatter && \
    chown -R dcformatter:dcformatter /app

# Switch to non-root user
USER dcformatter

# Health check to verify the application can be called
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python3 -m tools3.pipeline --help > /dev/null 2>&1 || exit 1

# Default command - show help
CMD ["python3", "-m", "tools3.pipeline", "--help"]
