# ─────────────────────────────────────────────────────────────────
#  Retail AI — Dockerfile
#  FastAPI + Python 3.11 slim
#  Ollama ayrı bir service olarak çalışmalıdır (sidecar pattern)
# ─────────────────────────────────────────────────────────────────

FROM python:3.11-slim

# System dependencies
RUN apt-get update && apt-get install -y \
    ffmpeg \
    && rm -rf /var/lib/apt/lists/*

# Working directory
WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project files
COPY . .

# Create data directory (for uploaded files)
RUN mkdir -p data

# Expose port
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:8000/ || exit 1

# Start server with dynamic Railway PORT
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}
