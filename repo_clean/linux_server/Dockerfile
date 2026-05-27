# ==========================================
# Dockerfile — JAGA TECH Report Engine V5 Linux
# Compatible with ARM64 (Oracle Linux / aarch64)
# ==========================================

FROM oraclelinux:9-slim

# Install system dependencies
RUN microdnf install -y \
    python3 \
    python3-pip \
    libreoffice-writer \
    libreoffice-headless \
    fontconfig \
    dejavu-sans-fonts \
    liberation-serif-fonts \
    liberation-sans-fonts \
    liberation-mono-fonts \
    && microdnf clean all

# Set working directory
WORKDIR /app

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create necessary directories
RUN mkdir -p uploads output

# Expose port
EXPOSE 8080

# Environment variables
ENV PORT=8080
ENV HOST=0.0.0.0
ENV PYTHONUNBUFFERED=1

# Health check
HEALTHCHECK --interval=30s --timeout=10s --retries=3 \
    CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:8080/health')" || exit 1

# Run with gunicorn for production
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "2", "--threads", "4", "--timeout", "300", "app_linux:app"]
