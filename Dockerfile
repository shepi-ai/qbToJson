# Use Python 3.11 slim image for smaller size
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies required for PDF processing
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    libpq-dev \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better layer caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY *.py ./
COPY converters/ ./converters/
COPY dialects/ ./dialects/
COPY renderers/ ./renderers/
COPY canonical/ ./canonical/
COPY sampleJson/ ./sampleJson/

# Create directory for temporary files
RUN mkdir -p /tmp/uploads

# Set environment variables
ENV PYTHONUNBUFFERED=1
ENV PORT=8080

# Expose port (Cloud Run will override this)
EXPOSE 8080

# Run the application with gunicorn for production
RUN pip install gunicorn

# Health check (simple curl-based)
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:8080/health || exit 1

# Use gunicorn with proper workers for Cloud Run
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 300 api_server:app
