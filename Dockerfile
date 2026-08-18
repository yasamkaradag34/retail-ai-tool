# ─────────────────────────────────────────────────────────────────
#  Retail AI — Dockerfile
#  FastAPI + Python 3.11 slim
#  Ollama ayrı bir service olarak çalışmalıdır (sidecar pattern)
# ─────────────────────────────────────────────────────────────────

FROM python:3.11-slim

# System dependencies
RUN apt-get update && apt-get install -y \
    curl \
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

# Start server on port 8000
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]

